using System.Collections.Concurrent;
using System.Runtime.CompilerServices;
using Invigoration.Sc2.Wire;

[assembly: InternalsVisibleTo("Invigoration.Sc2.Tests")]

namespace Invigoration.Sc2.Front;

/// <summary>A Front RPC answered with a non-zero status.</summary>
public sealed class FrontRpcException(string operation, uint status)
    : InvalidOperationException($"{operation} failed with status {status}.")
{
    public string Operation { get; } = operation;
    public uint Status { get; } = status;
}

/// <summary>
/// Battle.net account friends and presence over the Front connection, kept open after logon.
///
/// Up to logon (and optionally the Sunken handoff) FrontClient reads frames inline, one call at a
/// time. <see cref="StartListening"/> switches it to a background receive loop: from then on
/// responses are matched to calls by token, Echo keepalives are still answered, and
/// FriendsListener / PresenceListener / ChannelListener calls update <see cref="Social"/> and raise
/// events. Those listener methods are all declared NO_RESPONSE (as AuthenticationListener 11-13
/// are), so nothing is sent back for them; only Echo gets a reply.
///
/// Events are raised on the receive loop's thread; marshal to the UI yourself.
/// </summary>
public sealed partial class FrontClient
{
    private const uint MethodIdMask = 0x3FFFFFFF; // top two bits are client/server routing flags

    private static readonly uint FriendsServiceHash = ServiceHash.Compute(FrontServices.Friends);
    private static readonly uint PresenceServiceHash = ServiceHash.Compute(FrontServices.Presence);
    private static readonly uint[] FriendsListenerHashes =
        [ServiceHash.Compute(FrontServices.FriendsListener), ServiceHash.Compute(FrontServices.FriendsListenerV1)];
    private static readonly uint[] PresenceListenerHashes =
        [ServiceHash.Compute(FrontServices.PresenceListener), ServiceHash.Compute(FrontServices.PresenceListenerV1)];
    private static readonly uint[] ChannelListenerHashes =
        [ServiceHash.Compute(FrontServices.ChannelListener), ServiceHash.Compute(FrontServices.ChannelListenerV1)];

    /// <summary>Presence programs asked for when the caller names none: Battle.net's own fields ("BN"), which carry online, program and rich presence.</summary>
    public static readonly IReadOnlyList<string> DefaultPresencePrograms = ["BN"];

    private readonly ConcurrentDictionary<uint, (TaskCompletionSource<byte[]> Completion, string Operation)> _pending = new();
    private readonly ConcurrentDictionary<ulong, BgsEntityKey> _presenceObjects = new();
    private readonly object _subscribedGate = new();
    private readonly HashSet<BgsEntityKey> _presenceSubscribed = [];
    private readonly CancellationTokenSource _listenCts = new();
    private Task? _listenTask;
    private volatile bool _listenEnded;
    private volatile bool _closing;
    private volatile bool _followPresence;
    private IReadOnlyList<string> _followPrograms = DefaultPresencePrograms;
    private long _lastObjectId;

    /// <summary>The friends list and their presence, kept current once friends are subscribed.</summary>
    public BgsFriendsDirectory Social { get; } = new();

    public bool IsListening => _listenTask is not null && !_listenEnded;

    /// <summary>Any FriendsListener call (friend added/removed/updated, invitations).</summary>
    public event Action<BgsFriendsNotification>? FriendsNotified;

    /// <summary>Any presence change, from either PresenceListener or ChannelListener.</summary>
    public event Action<BgsPresenceUpdate>? PresenceUpdated;

    /// <summary>Raised after <see cref="Social"/> changed; re-read <see cref="BgsFriendsDirectory.GetFriends"/>.</summary>
    public event Action? SocialChanged;

    /// <summary>A background presence subscription failed, or an event handler threw.</summary>
    public event Action<Exception>? SocialError;

    /// <summary>A server call this client does not handle (header and body), for diagnostics.</summary>
    public event Action<Header, byte[]>? UnhandledCall;

    /// <summary>The receive loop ended: null after <see cref="CloseAsync"/>/dispose, otherwise why.</summary>
    public event Action<Exception?>? ListeningStopped;

    /// <summary>
    /// Starts the background receive loop. Call after <see cref="AuthenticateAsync"/> (and after
    /// <see cref="ProcessClientRequestAsync"/> if you do the handoff first; both calls also work
    /// after this). Idempotent.
    /// </summary>
    public void StartListening()
    {
        RequireConnected();
        if (_listenTask is not null)
        {
            return;
        }

        _listenTask = Task.Run(() => ListenLoopAsync(_listenCts.Token));
    }

    /// <summary>
    /// One call that does the whole thing: starts listening, subscribes to the friends list,
    /// subscribes to each friend account's presence, and (when <paramref name="followGameAccounts"/>)
    /// keeps following: game accounts revealed by account presence and friends added later are
    /// subscribed automatically. Returns the list as it stands once the subscriptions are answered.
    /// </summary>
    public async Task<IReadOnlyList<BgsFriendStatus>> StartSocialAsync(
        bool followGameAccounts = true,
        IReadOnlyList<string>? programs = null,
        CancellationToken cancellationToken = default)
    {
        _followPrograms = programs ?? DefaultPresencePrograms;
        _followPresence = followGameAccounts;
        StartListening();
        var friends = await SubscribeFriendsAsync(cancellationToken).ConfigureAwait(false);
        await SubscribePresenceAsync(friends.Friends.Select(f => f.AccountId), _followPrograms, cancellationToken).ConfigureAwait(false);
        return Social.GetFriends();
    }

    /// <summary>FriendsService.Subscribe: returns the current list (also applied to <see cref="Social"/>); later changes arrive through <see cref="FriendsNotified"/>.</summary>
    public async Task<BgsFriendsSubscribeResponse> SubscribeFriendsAsync(CancellationToken cancellationToken = default)
    {
        StartListening();
        var request = new BgsFriendsSubscribeRequest { ObjectId = NextObjectId() };
        var body = await CallAsync(FriendsServiceHash, 1, request.Encode(), "FriendsService.Subscribe", cancellationToken).ConfigureAwait(false);
        var response = BgsFriendsSubscribeResponse.Decode(body);
        Social.Apply(response);
        Raise(SocialChanged);
        return response;
    }

    /// <summary>
    /// PresenceService.Subscribe (method 1) for each entity not already subscribed, one object id
    /// per entity. Pass friend ACCOUNT ids (BN group 1: name, BattleTag, game-account list) and
    /// GAME ACCOUNT ids (BN group 2: online, program, away, rich presence). Throws an
    /// AggregateException naming any that Battle.net refused, after trying them all.
    /// </summary>
    public async Task SubscribePresenceAsync(
        IEnumerable<EntityId> entityIds,
        IReadOnlyList<string>? programs = null,
        CancellationToken cancellationToken = default)
    {
        StartListening();
        var programCodes = (programs ?? DefaultPresencePrograms).Select(FourCc.Encode).ToList();
        List<Task> calls = [];
        foreach (var entity in entityIds)
        {
            var key = BgsEntityKey.From(entity);
            lock (_subscribedGate)
            {
                if (!_presenceSubscribed.Add(key))
                {
                    continue;
                }
            }

            calls.Add(SubscribeOneAsync(entity, key, programCodes, cancellationToken));
        }

        try
        {
            await Task.WhenAll(calls).ConfigureAwait(false);
        }
        catch
        {
            var failures = calls.Where(c => c.IsFaulted).SelectMany(c => c.Exception!.InnerExceptions).ToList();
            if (failures.Count > 0)
            {
                throw new AggregateException("Some presence subscriptions failed.", failures);
            }

            throw;
        }
    }

    /// <summary>PresenceService.BatchSubscribe (method 8), the newer one-call form. Entities it reports as failed are left unsubscribed.</summary>
    public async Task<BgsPresenceBatchSubscribeResponse> BatchSubscribePresenceAsync(
        IEnumerable<EntityId> entityIds,
        IReadOnlyList<string>? programs = null,
        CancellationToken cancellationToken = default)
    {
        StartListening();
        var objectId = NextObjectId();
        var ids = entityIds.ToList();
        var request = new BgsPresenceBatchSubscribeRequest
        {
            EntityIds = ids,
            Programs = (programs ?? DefaultPresencePrograms).Select(FourCc.Encode).ToList(),
            ObjectId = objectId,
        };
        var body = await CallAsync(PresenceServiceHash, 8, request.Encode(), "PresenceService.BatchSubscribe", cancellationToken).ConfigureAwait(false);
        var response = BgsPresenceBatchSubscribeResponse.Decode(body);
        var failed = response.Failed.Where(f => f.EntityId is not null && (f.Result ?? 0) != 0)
            .Select(f => BgsEntityKey.From(f.EntityId!)).ToHashSet();
        lock (_subscribedGate)
        {
            foreach (var id in ids.Select(BgsEntityKey.From).Where(k => !failed.Contains(k)))
            {
                _presenceSubscribed.Add(id);
            }
        }

        return response;
    }

    /// <summary>PresenceService.Unsubscribe (method 2).</summary>
    public async Task UnsubscribePresenceAsync(EntityId entityId, CancellationToken cancellationToken = default)
    {
        var key = BgsEntityKey.From(entityId);
        var objectId = _presenceObjects.FirstOrDefault(p => p.Value == key).Key;
        var request = new BgsPresenceUnsubscribeRequest { EntityId = entityId, ObjectId = objectId == 0 ? null : objectId };
        await CallAsync(PresenceServiceHash, 2, request.Encode(), "PresenceService.Unsubscribe", cancellationToken).ConfigureAwait(false);
        lock (_subscribedGate)
        {
            _presenceSubscribed.Remove(key);
        }

        if (objectId != 0)
        {
            _presenceObjects.TryRemove(objectId, out _);
        }
    }

    private async Task SubscribeOneAsync(EntityId entity, BgsEntityKey key, List<uint> programs, CancellationToken cancellationToken)
    {
        var objectId = NextObjectId();
        _presenceObjects[objectId] = key;
        try
        {
            var request = new BgsPresenceSubscribeRequest { EntityId = entity, ObjectId = objectId, Programs = programs };
            await CallAsync(PresenceServiceHash, 1, request.Encode(), $"PresenceService.Subscribe({key})", cancellationToken).ConfigureAwait(false);
        }
        catch
        {
            _presenceObjects.TryRemove(objectId, out _);
            lock (_subscribedGate)
            {
                _presenceSubscribed.Remove(key);
            }

            throw;
        }
    }

    private ulong NextObjectId() => (ulong)Interlocked.Increment(ref _lastObjectId);

    private uint NextToken() => (uint)Interlocked.Increment(ref _lastToken);

    /// <summary>Sends a request and returns its response body, reading inline before <see cref="StartListening"/> and through the loop after.</summary>
    private async Task<byte[]> CallAsync(uint serviceHash, uint method, byte[] body, string operation, CancellationToken cancellationToken)
    {
        RequireConnected();
        if (_listenTask is null)
        {
            var inlineToken = await RequestAsync(serviceHash, method, body, cancellationToken).ConfigureAwait(false);
            var (_, response) = await AwaitResponseAsync(inlineToken, operation, cancellationToken).ConfigureAwait(false);
            return response;
        }

        if (_listenEnded)
        {
            throw new InvalidOperationException($"{operation}: the Front connection has stopped.");
        }

        var token = NextToken();
        var completion = new TaskCompletionSource<byte[]>(TaskCreationOptions.RunContinuationsAsynchronously);
        _pending[token] = (completion, operation);
        try
        {
            await SendAsync(new Header { ServiceId = 0, MethodId = method, Token = token, ServiceHash = serviceHash }, body, cancellationToken).ConfigureAwait(false);
            using (cancellationToken.Register(() => completion.TrySetCanceled(cancellationToken)))
            {
                return await completion.Task.ConfigureAwait(false);
            }
        }
        finally
        {
            _pending.TryRemove(token, out _);
        }
    }

    private async Task ListenLoopAsync(CancellationToken cancellationToken)
    {
        Exception? failure = null;
        try
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                var (header, body) = await ReceiveApplicationFrameAsync(cancellationToken).ConfigureAwait(false);
                Dispatch(header, body);
            }
        }
        catch (Exception ex)
        {
            if (!_closing)
            {
                failure = ex;
            }
        }
        finally
        {
            _listenEnded = true;
            var ended = failure ?? new InvalidOperationException("The Front connection was closed.");
            foreach (var token in _pending.Keys)
            {
                if (_pending.TryRemove(token, out var pending))
                {
                    pending.Completion.TrySetException(ended);
                }
            }

            try
            {
                ListeningStopped?.Invoke(failure);
            }
            catch
            {
                // Nothing left to report it to.
            }
        }
    }

    /// <summary>Handles one frame from the receive loop. Internal so tests can drive it without a socket.</summary>
    internal void Dispatch(Header header, byte[] body)
    {
        if (header.ServiceId == ResponseServiceId)
        {
            if (_pending.TryRemove(header.Token, out var pending))
            {
                if ((header.Status ?? 0) != 0)
                {
                    pending.Completion.TrySetException(new FrontRpcException(pending.Operation, header.Status!.Value));
                }
                else
                {
                    pending.Completion.TrySetResult(body);
                }
            }
            else
            {
                RaiseUnhandled(header, body);
            }

            return;
        }

        var hash = header.ServiceHash ?? 0;
        var method = (header.MethodId ?? 0) & MethodIdMask;
        try
        {
            if (FriendsListenerHashes.Contains(hash))
            {
                HandleFriends(method, header, body);
            }
            else if (PresenceListenerHashes.Contains(hash) && method is 1 or 2)
            {
                var notification = BgsPresenceListenerNotification.Decode(body);
                foreach (var state in notification.States)
                {
                    ApplyPresence(state, header.ObjectId, fullState: method == 1);
                }
            }
            else if (ChannelListenerHashes.Contains(hash) && BgsChannelPresence.Decode(method, body) is { } channelState)
            {
                ApplyPresence(channelState, header.ObjectId, fullState: method == BgsChannelPresence.OnJoinMethod);
            }
            else
            {
                RaiseUnhandled(header, body);
            }
        }
        catch (Exception ex) when (ex is IndexOutOfRangeException or ArgumentOutOfRangeException or InvalidOperationException)
        {
            // A body this client misreads must not take the connection down with it.
            RaiseError(new InvalidOperationException($"Could not read service_hash=0x{hash:X8} method={method}: {ex.Message}", ex));
        }
    }

    private void HandleFriends(uint method, Header header, byte[] body)
    {
        var notification = BgsFriendsNotification.Decode(method, body);
        if (notification is null)
        {
            RaiseUnhandled(header, body);
            return;
        }

        Social.Apply(notification);
        Raise(FriendsNotified, notification);
        Raise(SocialChanged);
        if (_followPresence && notification.Kind == BgsFriendsNotificationKind.FriendAdded && notification.Friend is not null)
        {
            Follow([notification.Friend.AccountId]);
        }
    }

    private void ApplyPresence(BgsPresenceState state, ulong? objectId, bool fullState)
    {
        BgsEntityKey entity;
        if (state.EntityId is not null)
        {
            entity = BgsEntityKey.From(state.EntityId);
        }
        else if (objectId is not null && _presenceObjects.TryGetValue(objectId.Value, out var mapped))
        {
            entity = mapped;
        }
        else
        {
            RaiseError(new InvalidOperationException($"Presence update for unknown object id {objectId}."));
            return;
        }

        var update = new BgsPresenceUpdate(entity, state.Operations, fullState || state.Healing == true);
        var discovered = Social.Apply(update);
        Raise(PresenceUpdated, update);
        Raise(SocialChanged);
        if (_followPresence && discovered.Count > 0)
        {
            Follow(discovered);
        }
    }

    /// <summary>Subscribes in the background: the loop cannot wait for a response it has to read itself.</summary>
    private void Follow(IReadOnlyList<EntityId> entities) =>
        _ = Task.Run(async () =>
        {
            try
            {
                await SubscribePresenceAsync(entities, _followPrograms, _listenCts.Token).ConfigureAwait(false);
            }
            catch (Exception ex) when (!_closing && !_listenEnded)
            {
                RaiseError(ex);
            }
            catch
            {
                // Closing: nothing to report.
            }
        });

    private void RequireNotListening(string operation)
    {
        if (_listenTask is not null)
        {
            throw new InvalidOperationException($"{operation} cannot run once the Front receive loop has started.");
        }
    }

    private void Raise(Action? handler)
    {
        try
        {
            handler?.Invoke();
        }
        catch (Exception ex)
        {
            RaiseError(ex);
        }
    }

    private void Raise<T>(Action<T>? handler, T value)
    {
        try
        {
            handler?.Invoke(value);
        }
        catch (Exception ex)
        {
            RaiseError(ex);
        }
    }

    private void RaiseUnhandled(Header header, byte[] body)
    {
        try
        {
            UnhandledCall?.Invoke(header, body);
        }
        catch (Exception ex)
        {
            RaiseError(ex);
        }
    }

    private void RaiseError(Exception ex)
    {
        try
        {
            SocialError?.Invoke(ex);
        }
        catch
        {
            // A throwing error handler has nowhere further to go.
        }
    }
}
