using System.Text;
using System.Threading.Channels;
using Invigoration.Core.Chat;
using FriendEntry = Invigoration.Core.Chat.FriendEntry;
using Invigoration.Core.Config;
using Sc2Friend = Invigoration.Sc2.Native.FriendEntry;
using Sc2FriendIdentity = Invigoration.Sc2.Native.FriendIdentity;
using WhisperTarget = Invigoration.Sc2.Chat.WhisperTarget;
using Invigoration.Sc2.Front;
using Invigoration.Sc2.Native;
using Invigoration.Sc2.Wire;
using Stimpak;

namespace Invigoration.Core.Sc2;

/// <summary>
/// Invigoration's own StarCraft II connection: Front sign-in, the handoff to Sunken, then chat on
/// Sunken's bit-packed records, with no Stimpak underneath. It reports in Stimpak's event types so
/// BotEngine treats it exactly like Stimpak's client. The default since 2.3.0 (<see cref="Enabled"/>);
/// it keeps a protocol log of every record it reads, with hex for anything it can't decode, so a
/// failed session says exactly where it stopped.
/// </summary>
/// <remarks>
/// Sign-in reuses the "keep me signed in" credential Battle.net issued last time (saved per game
/// under the bot's Battle.net profile). A web challenge only appears when Battle.net asks for one.
/// A new credential replaces the saved one only after it has been issued; nothing is deleted.
/// </remarks>
public sealed class NativeSc2ChatClient : ISc2ChatClient
{
    /// <summary>The program code this client signs in as.</summary>
    public const string Program = "S2";

    /// <summary>The old switch from when this client was test-only. Set to 0, it now falls back to Stimpak like <see cref="StimpakVariable"/>.</summary>
    public const string EnableVariable = "INVIGORATION_NATIVE_SC2";

    /// <summary>Set to 1 to use Stimpak's client instead of Invigoration's own, as a fallback. Native is the default from 2.3.0.</summary>
    public const string StimpakVariable = "INVIGORATION_STIMPAK";

    /// <summary>StarCraft II's public "General" channel.</summary>
    public const ushort GeneralChannelId = 1033;

    private static readonly Dictionary<ushort, string> KnownPublicChannels = new()
    {
        [1033] = "General",
        [1034] = "Trade",
        [1035] = "Help",
    };

    /// <summary>Whether SC2, SC:R and WC3:R bots use Invigoration's own connections (the default) rather than Stimpak's.</summary>
    public static bool Enabled =>
        Environment.GetEnvironmentVariable(StimpakVariable) != "1" && Environment.GetEnvironmentVariable(EnableVariable) != "0";

    private readonly string _profileId;
    private readonly Action<string> _log;
    private readonly Channel<SC2Event> _events = Channel.CreateUnbounded<SC2Event>(new UnboundedChannelOptions { SingleReader = true });
    private readonly SemaphoreSlim _sendLock = new(1, 1);
    private readonly object _sync = new();
    private readonly Dictionary<ulong, TaskCompletionSource<string>> _pendingAuth = new();
    private readonly Dictionary<uint, ChatChannel> _pendingJoins = new();
    private readonly Dictionary<byte, ChannelState> _channels = new();
    private CancellationTokenSource? _runCts;
    private Task? _run;
    private RecordStream? _stream;
    private ToonFullName? _self;
    private ulong _nextAuthId;
    private uint _nextJoinToken;
    private StreamWriter? _protocolLog;
    private bool _disposed;
    private byte[]? _presented;
    private readonly Dictionary<Sc2FriendIdentity, Sc2Friend> _friends = new();
    private PresenceTracker _presence = new();
    private readonly HashSet<uint> _askedFriendToons = new();
    private PortraitResolver _portraits = new();
    private int _presenceUpdatesLogged;
    private List<FriendEntry> _shownFriends = [];
    private IDisposable? _signInTurn;

    private sealed class ChannelState(ChatChannel channel, uint localHandle)
    {
        public ChatChannel Channel { get; } = channel;
        public uint LocalHandle { get; } = localHandle;
        public Dictionary<uint, User> Members { get; } = new();
        public bool RosterComplete { get; set; }
    }

    public NativeSc2ChatClient(string profileId, Action<string> log)
    {
        _profileId = profileId;
        _log = log;
        BattlenetFriendDirectory.Changed += OnFriendDirectoryChanged;
    }

    private void OnFriendDirectoryChanged(string profileId)
    {
        if (profileId == _profileId && _stream is not null)
        {
            EmitFriendsIfChanged();
        }
    }

    public PeopleRegistry People { get; } = new();

    public event Action<SC2Event>? EventReceived;

    public IAsyncEnumerable<SC2Event> ReadEventsAsync(CancellationToken cancellation = default) => _events.Reader.ReadAllAsync(cancellation);

    public void Connect(StimpakConnectOptions options)
    {
        lock (_sync)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_run is { IsCompleted: false })
            {
                throw new StimpakException("The native StarCraft II connection is already running.", null!);
            }

            _runCts = new CancellationTokenSource();
            var channels = options.Channels is { Count: > 0 } saved ? saved.ToList() : [ChannelTarget.Public(GeneralChannelId)];
            _run = Task.Run(() => RunAsync(channels, _runCts.Token));
        }
    }

    public void Disconnect()
    {
        lock (_sync)
        {
            _runCts?.Cancel();
        }
    }

    public void JoinPublic(ushort channelId) =>
        Send(ChatCommands.ChatJoinPublic(channelId, NewJoinToken(PublicChannelFor(channelId)), "enUS"), "join public channel");

    public void JoinPrivate(string name) =>
        Send(ChatCommands.ChatJoinPrivate(name, NewJoinToken(new PrivateChannel(name))), "join " + name);

    public void Leave(byte channelIndex) => Send(ChatCommands.ChatLeave(channelIndex), "leave channel");

    /// <summary>
    /// Sends to a channel, then reports the message back as received, from us: Battle.net doesn't
    /// echo a sender's own messages, and BotEngine (like Stimpak's client) shows sent lines only
    /// through that echo.
    /// </summary>
    public void SendMessage(byte channelIndex, string body)
    {
        Send(ChatCommands.ChatMessage(channelIndex, body), "send message");
        User self;
        lock (_sync)
        {
            var handle = _channels.TryGetValue(channelIndex, out var state) ? state.LocalHandle : 0;
            self = new User(handle, null, _self?.Name ?? "", "", Presence.Available);
        }

        Emit(new MessageReceived(channelIndex, self, body));
    }

    /// <summary>
    /// Whispers a character by name. The target's region and realm are taken from someone of
    /// that name in a joined channel if there is one, otherwise from our own character, which is
    /// the usual case: whispers stay within one region.
    /// </summary>
    public void SendWhisper(string name, string body)
    {
        // A friend by the name the Friends list shows: addressed by their presence, or their
        // account, as ncarrillo/superiority (MIT) does, so their SC2 character's name isn't needed.
        if (FriendWhisperTarget(name) is { } friend)
        {
            Send(ChatCommands.ChatWhisper(friend, body), "whisper " + name);
            Emit(new WhisperReceived(name, body, true));
            return;
        }

        var target = FindMember(name) ?? _self
            ?? throw new StimpakException("Can't whisper before a character is selected.", null!);
        Send(ChatCommands.ChatWhisper(new WhisperTarget.ToonName(name, target.Region, target.ProgramId, target.Realm), body), "whisper " + name);
        Emit(new WhisperReceived(name, body, true));
    }

    /// <summary>How to whisper the friend the Friends list shows as <paramref name="name"/>, or null for someone who isn't one.</summary>
    private WhisperTarget? FriendWhisperTarget(string name)
    {
        lock (_sync)
        {
            foreach (var friend in _friends.Values)
            {
                if (friend.Identity is not Sc2FriendIdentity.Account account
                    || !ToFriendEntry(friend).Account.Equals(name.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                return _presence.PresenceIdFor(friend) is { } presenceId
                    ? new WhisperTarget.Presence(presenceId)
                    : new WhisperTarget.Account(account.AccountId);
            }
        }

        return null;
    }

    public void SubmitAuth(ulong authId, string token)
    {
        lock (_sync)
        {
            if (_pendingAuth.Remove(authId, out var pending))
            {
                pending.TrySetResult(token);
            }
        }
    }

    public void CancelAuth(ulong authId)
    {
        lock (_sync)
        {
            if (_pendingAuth.Remove(authId, out var pending))
            {
                pending.TrySetCanceled();
            }
        }
    }

    public void Dispose()
    {
        lock (_sync)
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _runCts?.Cancel();
        }

        BattlenetFriendDirectory.Changed -= OnFriendDirectoryChanged;

        _ = Task.Run(async () =>
        {
            if (_run is { } run)
            {
                await run.ConfigureAwait(false);
            }

            _events.Writer.TryComplete();
            _sendLock.Dispose();
        });
    }

    private async Task RunAsync(IReadOnlyList<ChannelTarget> channels, CancellationToken cancellationToken)
    {
        OpenProtocolLog();
        try
        {
            while (true)
            {
                try
                {
                    await SignInAndChatAsync(channels, cancellationToken).ConfigureAwait(false);
                    break;
                }
                catch (SavedSignInArrivedException arrived)
                {
                    // Not a failure: another game on the profile signed in for us meanwhile.
                    Trace(arrived.Message);
                }
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            Trace("Disconnected on request.");
        }
        catch (Exception ex)
        {
            Trace($"Session failed: {ex}");
            if (_stream is { } stream)
            {
                Trace($"Unread bytes at failure: {stream.PendingHex()}");
            }

            Emit(new SessionFailed(ex.Message));
        }
        finally
        {
            EndSignInTurn();
            _stream?.Dispose();
            _stream = null;
            lock (_sync)
            {
                foreach (var pending in _pendingAuth.Values)
                {
                    pending.TrySetCanceled();
                }

                _pendingAuth.Clear();
                _pendingJoins.Clear();
                _channels.Clear();
            }

            Emit(new StageChanged(Stage.Disconnected));
            _protocolLog?.Dispose();
            _protocolLog = null;
        }
    }

    private async Task SignInAndChatAsync(IReadOnlyList<ChannelTarget> channels, CancellationToken cancellationToken)
    {
        lock (_sync)
        {
            _friends.Clear();
            _presence = new PresenceTracker();
            _portraits = new PortraitResolver();
            _presenceUpdatesLogged = 0;
            _shownFriends = [];
            _askedFriendToons.Clear();
        }

        Emit(new StageChanged(Stage.WebAuthentication));
        var front = new FrontClient();
        var frontOpen = true;
        try
        {
            Trace($"Opening Front at {FrontClient.DefaultUsUri}");
            await front.ConnectAsync(new Uri(FrontClient.DefaultUsUri), cancellationToken).ConfigureAwait(false);
            await front.EstablishAsync(cancellationToken).ConfigureAwait(false);

            var saved = _presented = BattlenetCredentialProfileStore.LoadNativeCredential(_profileId, Program);
            Trace(saved is null ? "No saved sign-in; Battle.net will ask for one." : "Presenting the saved sign-in.");
            var logon = await front.AuthenticateAsync(saved, ChallengeAsync, cancellationToken).ConfigureAwait(false);
            if (logon.ErrorCode != 0)
            {
                throw new StimpakException($"Battle.net sign-in failed (error {logon.ErrorCode}).", null!);
            }

            Trace($"Signed in as {logon.BattleTag}; {logon.GameAccountIds.Count} game account(s), region {logon.ConnectedRegion}.");
            Emit(new AccountConnected(new AccountSummary(logon.AccountId?.Low, logon.BattleTag, logon.ConnectedRegion, [])));

            // Replace the saved sign-in only now that a new one has actually been issued.
            var fresh = await front.GenerateWebCredentialsAsync(Program, cancellationToken).ConfigureAwait(false);
            BattlenetCredentialProfileStore.SaveNativeCredential(_profileId, Program, fresh);
            Trace("Saved the new sign-in.");
            await NativeSignIns.IssueMissingAsync(
                _profileId,
                Program,
                async (program, token) => await front.GenerateWebCredentialsAsync(program, token).ConfigureAwait(false),
                Trace,
                cancellationToken).ConfigureAwait(false);
            EndSignInTurn();

            Emit(new StageChanged(Stage.GameUtilities));
            var gameAccount = logon.GameAccountIds.FirstOrDefault()
                ?? throw new StimpakException("This Battle.net account has no StarCraft II game account.", null!);
            var sessionKey = logon.SessionKey ?? throw new StimpakException("Battle.net sent no session key.", null!);
            var handoff = await front.ProcessClientRequestAsync(gameAccount, sessionKey, cancellationToken).ConfigureAwait(false);
            Trace($"Handoff to Sunken at {handoff.Address}.");

            Emit(new StageChanged(Stage.NativeAuthentication));
            // Front closes once Sunken is reached. Keeping it for the account friends list was tried
            // (FrontClient.StartSocialAsync): Battle.net refuses FriendsService.Subscribe from a game
            // client with 3025 (ERROR_RPC_METHOD_DISABLED) and drops Front.
            var session = await SunkenClient.ConnectAsync(
                handoff,
                onStage: Trace,
                cancellationToken: cancellationToken,
                afterTcpConnect: async () =>
                {
                    frontOpen = false;
                    await front.CloseAsync(cancellationToken).ConfigureAwait(false);
                    Trace("Front closed.");
                }).ConfigureAwait(false);
            _stream = session.Stream;
            Trace($"Resumed: {session.Details.FinalRequests.Count} final request module(s), ping timeout {session.Details.PingTimeoutSeconds}s.");
        }
        finally
        {
            if (frontOpen)
            {
                try
                {
                    await front.CloseAsync(CancellationToken.None).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    Trace($"Closing Front: {ex.Message}");
                }
            }

            await front.DisposeAsync().ConfigureAwait(false);
        }

        Emit(new StageChanged(Stage.ChatBootstrap));
        using var pinger = new PeriodicTimer(ConnectionCommands.PingInterval);
        var pinging = PingLoopAsync(pinger, cancellationToken);
        using var portraitsStop = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var portraits = PumpPortraitsAsync(portraitsStop.Token);
        try
        {
            await ReadLoopAsync(channels, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            pinger.Dispose();
            await pinging.ConfigureAwait(ConfigureAwaitOptions.SuppressThrowing);
            await portraitsStop.CancelAsync().ConfigureAwait(false);
            await portraits.ConfigureAwait(ConfigureAwaitOptions.SuppressThrowing);
        }
    }

    private async Task ReadLoopAsync(IReadOnlyList<ChannelTarget> channels, CancellationToken cancellationToken)
    {
        var stream = _stream!;
        var joinedAny = false;
        var toonSelected = false;

        // Whether Battle.net sends the character list unprompted isn't confirmed yet. If it
        // doesn't, say so plainly in the log rather than sitting silent.
        _ = Task.Delay(TimeSpan.FromSeconds(15), cancellationToken).ContinueWith(
            _ =>
            {
                if (!toonSelected)
                {
                    Trace("Still no character list 15 seconds after resuming. Battle.net may be waiting for a request first.");
                }
            },
            TaskContinuationOptions.OnlyOnRanToCompletion);
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var before = stream.PendingHex();
            bool decoded;
            NativeChatRecord? record;
            try
            {
                decoded = stream.TryDecodeRecord(NativeRecordDispatcher.Decode, out record);
            }
            catch (UnknownNativeRecordException unknown)
            {
                // No layout for this record, so no length. It most likely arrived on its own, so
                // dropping what's buffered skips exactly it; if not, the next record won't decode and
                // the session ends as it would have anyway. Its bytes are kept here for a decoder.
                var dropped = stream.DiscardPending();
                Trace($"Unknown record slot {unknown.Slot} command {unknown.Command}; skipped the {dropped} buffered byte(s): {before}");
                continue;
            }

            if (!decoded)
            {
                if (!await stream.FillAsync(cancellationToken).ConfigureAwait(false))
                {
                    throw new IOException("Battle.net closed the StarCraft II connection.");
                }

                continue;
            }

            TraceRecord(record!, before, stream.PendingHex());
            switch (record)
            {
                case NativeChatRecord.Ping ping:
                    await SendAsync(ConnectionCommands.Pong(ping.Timestamp), cancellationToken).ConfigureAwait(false);
                    break;

                case NativeChatRecord.ToonList list when !toonSelected:
                    var toon = list.Value.Displays.FirstOrDefault()
                        ?? throw new StimpakException("This account has no StarCraft II character.", null!);
                    Trace($"Selecting character {toon.Name} (realm {toon.Realm}).");
                    await SendAsync(ChatCommands.ToonSelect(toon.Name, toon.Realm), cancellationToken).ConfigureAwait(false);
                    toonSelected = true;
                    break;

                case NativeChatRecord.ToonSelected selected:
                    _self = new ToonFullName(selected.Value.Handle.Region, selected.Value.Handle.Program, selected.Value.Realm, selected.Value.ToonName);
                    Trace($"Character selected: {_self.Name}. Joining {channels.Count} channel(s).");

                    // The "join another channel" picker lists these. Battle.net's own channel-list
                    // reply carries IDs, not names, so the known public channels are offered by name.
                    Emit(new PublicChannelsReceived(KnownPublicChannels.Select(c => (ChatChannel)new PublicChannel(c.Key, c.Value)).ToList()));
                    Send(ChatCommands.ChatChannelListRequest(), "ask for the public channel list");
                    foreach (var target in channels)
                    {
                        JoinTarget(target);
                    }

                    break;

                case NativeChatRecord.Join join:
                    joinedAny |= HandleJoin(join.Value);
                    if (joinedAny && _channels.Count == 1 && join.Value.Success)
                    {
                        Emit(new StageChanged(Stage.Connected));
                    }

                    break;

                case NativeChatRecord.PublicChannelList list:
                    Trace($"Public channel list: {string.Join("; ", list.Entries.Select(e => $"{e.A}/{e.B}/{e.C}"))}");
                    break;

                case NativeChatRecord.Membership membership:
                    HandleMembership(membership.Value);
                    UpdatePortraits();
                    break;

                case NativeChatRecord.Message message:
                    HandleMessage(message.Value);
                    break;

                case NativeChatRecord.FriendsList friends:
                    HandleFriends(friends.Value);
                    break;

                case NativeChatRecord.ToonsOfFriends toons:
                    HandleFriendToons(toons.Value);
                    break;

                case NativeChatRecord.PresenceFields fields:
                    _presence.Announce(fields.Value);
                    _protocolLog?.WriteLine($"{DateTime.Now:HH:mm:ss.fff}    presence fields: "
                        + string.Join(", ", fields.Value.Fields.Select(f => $"{f.Handle:X}:t{f.TypeId}{(f.FixedSize is { } size ? $"/{size}" : "")}")));
                    break;

                case NativeChatRecord.PresenceUpdate update:
                    // Every value, for the first few hundred updates: to find whether SC2's presence
                    // says which game (not just SC2) a friend is in.
                    if (_presenceUpdatesLogged++ < 300)
                    {
                        var u = update.Value;
                        _protocolLog?.WriteLine($"{DateTime.Now:HH:mm:ss.fff}    presence {u.MasterPresenceId}/{u.LocalPresenceId} online={u.Online} "
                            + $"fields [{string.Join(",", u.Handles.Select(h => h.ToString("X")))}] cleared [{string.Join(",", u.ClearedHandles.Select(h => h.ToString("X")))}] "
                            + $"sizes [{string.Join(",", u.VariableSizes)}] data {Convert.ToHexString(u.FieldData.AsSpan(0, Math.Min(u.FieldData.Length, 96)))}");
                    }

                    bool applied;
                    lock (_sync)
                    {
                        applied = _presence.Apply(update.Value);
                    }

                    if (applied)
                    {
                        EmitFriendsIfChanged();
                        UpdatePortraits();
                    }

                    break;

                case NativeChatRecord.ProfileRead profile:
                    bool resolved;
                    lock (_sync)
                    {
                        resolved = _portraits.Complete(profile.Value);
                    }

                    if (resolved)
                    {
                        UpdatePortraits();
                    }

                    break;

                case NativeChatRecord.Whisper whisper:
                    Emit(new WhisperReceived(whisper.Value.PeerName, whisper.Value.Body, false));
                    break;
            }
        }
    }

    /// <summary>
    /// Channel members' profile portraits, as superiority (MIT) finds them: straight from presence
    /// when it carries one, otherwise by reading each member's profile, rate-limited by
    /// PortraitResolver (see PumpPortraitsAsync). Found ones go to NativeMemberPortraits for the user list.
    /// </summary>
    private void UpdatePortraits()
    {
        lock (_sync)
        {
            foreach (var user in _channels.Values.SelectMany(c => c.Members.Values))
            {
                if (user.PresenceId is not { } presenceId)
                {
                    continue;
                }

                if (_portraits.For(_presence, presenceId) is { } portrait)
                {
                    NativeMemberPortraits.Set(user.Name, portrait.Table, portrait.Offset);
                }
                else
                {
                    _portraits.EnqueueMember(_presence, presenceId);
                }

                NativeMemberPortraits.SetDetail(user.Name, MemberDetail(presenceId));
                if (_presence.State(presenceId) is { } state)
                {
                    NativeMemberPortraits.SetState(user.Name, state switch
                    {
                        FriendPresence.Away => Presence.Away,
                        FriendPresence.Busy => Presence.Busy,
                        FriendPresence.InGame => Presence.InGame,
                        FriendPresence.Offline => Presence.Offline,
                        _ => Presence.Available,
                    });
                }
            }
        }
    }

    /// <summary>What presence says about a member, for the user list's Full view: BattleTag, and in a game, away or busy.</summary>
    private string MemberDetail(uint presenceId)
    {
        var parts = new List<string>();
        if (_presence.BattleTagFor(presenceId) is { Length: > 0 } tag)
        {
            parts.Add(tag);
        }

        switch (_presence.State(presenceId))
        {
            case FriendPresence.InGame:
                parts.Add("in a game");
                break;
            case FriendPresence.Away:
                parts.Add("away");
                break;
            case FriendPresence.Busy:
                parts.Add("busy");
                break;
        }

        return string.Join(" · ", parts);
    }

    /// <summary>Sends the profile reads PortraitResolver allows (16 at once, 40 a second), four times a second.</summary>
    private async Task PumpPortraitsAsync(CancellationToken cancellationToken)
    {
        using var timer = new PeriodicTimer(TimeSpan.FromMilliseconds(250));
        try
        {
            while (await timer.WaitForNextTickAsync(cancellationToken).ConfigureAwait(false))
            {
                List<(uint RequestId, PlayerTarget.ProfileRecordAddress Address)> requests;
                lock (_sync)
                {
                    requests = _portraits.NextRequests(DateTimeOffset.UtcNow).ToList();
                }

                foreach (var (requestId, address) in requests)
                {
                    await SendAsync(ChatCommands.ProfileReadRequest(requestId, address), cancellationToken).ConfigureAwait(false);
                }
            }
        }
        catch (Exception ex) when (ex is OperationCanceledException or ObjectDisposedException)
        {
            // Session over.
        }
        catch (Exception ex)
        {
            Trace($"Portrait requests stopped: {ex.Message}");
        }
    }

    private bool HandleJoin(ChatJoinRecord join)
    {
        ChatChannel? requested = null;
        lock (_sync)
        {
            if (join.Token is { } token && _pendingJoins.Remove(token, out var pending))
            {
                requested = pending;
            }
        }

        if (!join.Success || join.ChannelIndex is not { } index)
        {
            Trace($"Join refused by Battle.net: {requested?.Name ?? "unknown channel"}, reason {join.Reason}.");
            Emit(new JoinRejected(requested ?? new PrivateChannel("?"), join.Reason));
            return false;
        }

        var channel = requested
            ?? (join.ChannelNameId is { } nameId ? PublicChannelFor(nameId) : new PrivateChannel($"Channel {index}"));
        lock (_sync)
        {
            _channels[index] = new ChannelState(channel, join.MemberHandle ?? 0);
        }

        Emit(new Joined(index, channel, join.MemberHandle ?? 0));
        return true;
    }

    /// <summary>
    /// FriendsListNotify5 (Friends slot, command 30): Battle.net pushes the list at startup in pages
    /// and sends changes later. Each friend is shown by BattleTag (or SC2 character); the real name
    /// the record also carries is never shown.
    /// </summary>
    private void HandleFriends(FriendsListRecord page)
    {
        lock (_sync)
        {
            foreach (var update in page.Updates)
            {
                var id = update.Entry.Identity;
                switch (update.Operation)
                {
                    case SocialOperation.Remove:
                        _friends.Remove(id);
                        break;
                    case SocialOperation.Modify when _friends.TryGetValue(id, out var known):
                        _friends[id] = known with
                        {
                            DisplayName = update.Entry.DisplayName ?? known.DisplayName,
                            ToonName = update.Entry.ToonName ?? known.ToonName,
                            Profile = update.Entry.Profile ?? known.Profile,
                        };
                        break;
                    default:
                        _friends[id] = update.Entry;
                        break;
                }
            }
        }

        EmitFriendsIfChanged(page.Complete == false ? ", more to come" : "");

        // Friends sent without a name: ask for their SC2 characters, as the game (and Stimpak) do.
        List<uint> ask;
        lock (_sync)
        {
            ask = _friends.Values
                .Where(f => f.DisplayName is not { Length: > 0 } && f.ToonName is null)
                .Select(f => f.Identity).OfType<Sc2FriendIdentity.Account>().Select(a => a.AccountId)
                .Where(_askedFriendToons.Add)
                .ToList();
        }

        foreach (var accountId in ask)
        {
            Send(ChatCommands.ToonsOfFriendsRequest(accountId), "ask for a friend's characters");
        }
    }

    /// <summary>ToonsOfFriendsNotify: a friend's SC2 characters, which is the name they're shown by.</summary>
    private void HandleFriendToons(ToonsOfFriendsRecord record)
    {
        lock (_sync)
        {
            foreach (var toon in record.Entries)
            {
                var id = new Sc2FriendIdentity.Account(toon.AccountId);
                if (_friends.TryGetValue(id, out var friend) && friend.ToonName is null)
                {
                    _friends[id] = friend with { ToonName = toon.ToonName, Profile = friend.Profile ?? toon.Profile };
                }
            }
        }

        EmitFriendsIfChanged();
    }

    /// <summary>Reports the friends list when it or anyone's status has changed; presence updates arrive often.</summary>
    private void EmitFriendsIfChanged(string note = "")
    {
        List<FriendEntry> list;
        lock (_sync)
        {
            list = _friends.Values.Select(ToFriendEntry)
                .OrderBy(f => f.Location == FriendLocation.Offline).ThenBy(f => f.Account, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (list.SequenceEqual(_shownFriends))
            {
                return;
            }

            _shownFriends = list;
        }

        int fromDirectory;
        lock (_sync)
        {
            fromDirectory = _friends.Keys.OfType<Sc2FriendIdentity.Account>().Count(a => BattlenetFriendDirectory.Find(_profileId, a.AccountId) is not null);
        }

        Trace($"Friends list: {list.Count} ({list.Count(f => f.Location != FriendLocation.Offline)} online){note}; {fromDirectory} named from the Battle.net friends list.");
        Emit(new NativeFriendsEvent(list));
    }

    /// <summary>
    /// A friend as the Battle.net app shows them. SC2 doesn't send BattleTags for account friends,
    /// so those come from the profile's Battle.net friends list as SC:R received it
    /// (BattlenetFriendDirectory), and so does what they're doing while an SC:R bot is connected.
    /// Otherwise: their SC2 character, else their real name, else their ID, and SC2's own presence.
    /// </summary>
    private FriendEntry ToFriendEntry(Sc2Friend friend)
    {
        var accountId = (friend.Identity as Sc2FriendIdentity.Account)?.AccountId;
        var known = accountId is { } id ? BattlenetFriendDirectory.Find(_profileId, id) : null;
        var realName = known?.RealName is { Length: > 0 } shared ? shared : friend.FullName ?? "";
        var name = known?.BattleTag
            ?? _presence.BattleTagFor(friend)
            ?? (friend.DisplayName is { Length: > 0 } tag ? tag
                : friend.ToonName is { } toon ? toon.Name
                : realName.Length > 0 ? realName
                : accountId is { } number ? $"#{number}" : "?");
        if (name == realName)
        {
            realName = "";
        }

        if (accountId is { } liveId && BattlenetFriendDirectory.ActivityOf(_profileId, liveId) is { } activity)
        {
            if (!activity.Online)
            {
                return new FriendEntry(name, FriendStatus.None, FriendLocation.Offline, "", "", realName);
            }

            var (icon, game) = BattlenetPrograms.Describe(activity.Program);
            return new FriendEntry(name, FriendStatus.None, FriendLocation.NotInChat, icon,
                activity.Detail.Length > 0 ? $"{game}: {activity.Detail}" : game, realName);
        }

        // SC2's presence names the game account a friend is on (field 0x10018), as SC:R's list does.
        var state = _presence.For(friend);
        if (state is null or FriendPresence.Offline)
        {
            return new FriendEntry(name, FriendStatus.None, FriendLocation.Offline, "", "", realName);
        }

        var (gameIcon, gameName) = BattlenetPrograms.Describe(_presence.ProgramFor(friend) ?? "");
        var (status, doing) = state switch
        {
            FriendPresence.Away => (FriendStatus.Away, $"{gameName} (away)"),
            FriendPresence.Busy => (FriendStatus.DoNotDisturb, $"{gameName} (busy)"),
            FriendPresence.InGame => (FriendStatus.None, $"{gameName}: in a game"),
            _ => (FriendStatus.None, gameName),
        };
        return new FriendEntry(name, status, FriendLocation.NotInChat, gameIcon, doing, realName);
    }

    private void HandleMembership(MembershipChangeNotifyRecord record)
    {
        ChannelState? state;
        lock (_sync)
        {
            _channels.TryGetValue(record.ChannelIndex, out state);
        }

        if (state is null)
        {
            Trace($"Membership change for channel {record.ChannelIndex}, which isn't joined; ignored.");
            return;
        }

        var initial = !state.RosterComplete;
        var snapshot = new List<User>();
        foreach (var change in record.Changes)
        {
            switch (change)
            {
                case MembershipChange.Join join:
                    var user = new User(join.MemberHandle, join.PresenceId, DisplayName(join.Statuses) ?? $"#{join.MemberHandle}", "", Presence.Available);
                    state.Members[join.MemberHandle] = user;
                    if (initial)
                    {
                        snapshot.Add(user);
                    }
                    else
                    {
                        Emit(new MemberJoined(record.ChannelIndex, user));
                    }

                    break;

                case MembershipChange.Update { Status: MemberStatus.Display display } update:
                    if (state.Members.TryGetValue(update.MemberHandle, out var known))
                    {
                        state.Members[update.MemberHandle] = known with { Name = display.ToonName.Name };
                    }

                    break;

                case MembershipChange.Leave leave:
                    if (state.Members.Remove(leave.MemberHandle, out var left))
                    {
                        Emit(new MemberLeft(record.ChannelIndex, left));
                    }

                    break;
            }
        }

        if (initial)
        {
            state.RosterComplete = record.EndOfInitial;
            Emit(new RosterReceived(record.ChannelIndex, record.EndOfInitial, snapshot));
        }
    }

    private void HandleMessage(ChatMessageRecord message)
    {
        User? sender = null;
        lock (_sync)
        {
            if (_channels.TryGetValue(message.ChannelIndex, out var state))
            {
                state.Members.TryGetValue(message.MemberHandle, out sender);
            }
        }

        Emit(new MessageReceived(message.ChannelIndex, sender ?? new User(message.MemberHandle, null, $"#{message.MemberHandle}", "", Presence.Available), message.Body));
    }

    private static string? DisplayName(IEnumerable<MemberStatus> statuses) =>
        statuses.OfType<MemberStatus.Display>().LastOrDefault()?.ToonName.Name;

    private ToonFullName? FindMember(string name)
    {
        lock (_sync)
        {
            // Only the name is known for members; region and realm come from our own character.
            var found = _channels.Values.SelectMany(c => c.Members.Values).Any(u => string.Equals(u.Name, name, StringComparison.OrdinalIgnoreCase));
            return found ? _self : null;
        }
    }

    private void JoinTarget(ChannelTarget target)
    {
        switch (target)
        {
            case PublicChannelTarget pub:
                JoinPublic(pub.Id);
                break;
            case PrivateChannelTarget priv:
                JoinPrivate(priv.Name);
                break;
            case GroupChannelTarget group:
                Trace($"Skipping group channel {group.ClubId}: groups aren't supported by the native connection yet.");
                break;
        }
    }

    private static PublicChannel PublicChannelFor(ushort id) =>
        new(id, KnownPublicChannels.TryGetValue(id, out var name) ? name : $"Channel {id}");

    private uint NewJoinToken(ChatChannel channel)
    {
        lock (_sync)
        {
            var token = ++_nextJoinToken;
            _pendingJoins[token] = channel;
            return token;
        }
    }

    private void EndSignInTurn() => Interlocked.Exchange(ref _signInTurn, null)?.Dispose();

    private async Task<byte[]> ChallengeAsync(Uri url, CancellationToken cancellationToken)
    {
        // One sign-in window per profile at a time; kept until the other games' sign-ins are saved.
        _signInTurn ??= await NativeSignIns.BeginWebSignInAsync(_profileId, Program, _presented, cancellationToken).ConfigureAwait(false);

        TaskCompletionSource<string> answer;
        ulong authId;
        lock (_sync)
        {
            authId = ++_nextAuthId;
            answer = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
            _pendingAuth[authId] = answer;
        }

        // Where the sign-in page is, without values: query values can carry sign-in tokens.
        var parameters = System.Web.HttpUtility.ParseQueryString(url.Query).AllKeys;
        Trace($"Battle.net asked for a web sign-in at {url.Scheme}://{url.Host}{url.AbsolutePath} (parameters: {string.Join(", ", parameters)}).");
        Emit(new AuthenticationRequired(authId, url.ToString(), false));
        using var registration = cancellationToken.Register(() => answer.TrySetCanceled(cancellationToken));
        var token = await answer.Task.ConfigureAwait(false);
        return Encoding.UTF8.GetBytes(token);
    }

    private async Task PingLoopAsync(PeriodicTimer timer, CancellationToken cancellationToken)
    {
        try
        {
            while (await timer.WaitForNextTickAsync(cancellationToken).ConfigureAwait(false))
            {
                var micros = (ulong)(DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() * 1000);
                await SendAsync(ConnectionCommands.Ping(micros), cancellationToken).ConfigureAwait(false);
            }
        }
        catch (Exception ex) when (ex is OperationCanceledException or ObjectDisposedException)
        {
            // Session over.
        }
    }

    private void Send(byte[] record, string what)
    {
        if (_stream is null)
        {
            throw new StimpakException($"Can't {what}: not connected to StarCraft II chat.", null!);
        }

        _ = SendReportingAsync(record, what);
    }

    private async Task SendReportingAsync(byte[] record, string what)
    {
        try
        {
            await SendAsync(record, _runCts?.Token ?? CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            Trace($"Couldn't {what}: {ex.Message}");
            Emit(new CommandFailed($"Couldn't {what}: {ex.Message}"));
        }
    }

    private async Task SendAsync(byte[] record, CancellationToken cancellationToken)
    {
        var stream = _stream ?? throw new InvalidOperationException("Not connected.");
        await _sendLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            TraceOutgoing(record);
            await stream.SendAsync(record, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _sendLock.Release();
        }
    }

    private void Emit(SC2Event next)
    {
        EventReceived?.Invoke(next);
        _events.Writer.TryWrite(next);
    }

    private void OpenProtocolLog()
    {
        try
        {
            var directory = Path.Combine(ConfigStore.DefaultConfigDirectory(), "Logs");
            Directory.CreateDirectory(directory);
            var path = Path.Combine(directory, $"sc2-native-{DateTime.Now:yyyyMMdd-HHmmss}.log");
            _protocolLog = new StreamWriter(path, append: false, Encoding.UTF8) { AutoFlush = true };
            _log($"Native StarCraft II protocol log: {path}");
        }
        catch (Exception ex)
        {
            _log($"Couldn't open the native StarCraft II protocol log: {ex.Message}");
        }
    }

    /// <summary>A line for both the bot's debug log and the protocol log. Never given credentials or keys.</summary>
    private void Trace(string message)
    {
        _log($"[native SC2] {message}");
        _protocolLog?.WriteLine($"{DateTime.Now:HH:mm:ss.fff} {message}");
    }

    private void TraceRecord(NativeChatRecord record, string before, string after)
    {
        var consumedHex = before.Length >= after.Length ? before[..(before.Length - after.Length)] : before;
        var detail = record switch
        {
            NativeChatRecord.Sc2Consumed consumed => $"consumed slot {consumed.Slot} command {consumed.Command}",
            NativeChatRecord.Message => "chat message",
            _ => record.GetType().Name,
        };
        _protocolLog?.WriteLine($"{DateTime.Now:HH:mm:ss.fff} <- {detail} ({consumedHex.Length / 2} bytes) {consumedHex}");
    }

    private void TraceOutgoing(byte[] record)
    {
        var reader = new BitReader(record);
        var route = RoutingHeader.Decode(reader);
        _protocolLog?.WriteLine($"{DateTime.Now:HH:mm:ss.fff} -> slot {route.ServiceSlot} command {route.CommandId} ({record.Length} bytes)");
    }
}
