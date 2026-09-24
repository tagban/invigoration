using System.Text;
using System.Threading.Channels;
using Invigoration.Core.Config;
using WhisperTarget = Invigoration.Sc2.Chat.WhisperTarget;
using Invigoration.Sc2.Front;
using Invigoration.Sc2.Native;
using Invigoration.Sc2.Wire;
using Stimpak;

namespace Invigoration.Core.Sc2;

/// <summary>
/// Invigoration's own StarCraft II connection: Front sign-in, the handoff to Sunken, then chat on
/// Sunken's bit-packed records, with no Stimpak underneath. It reports in Stimpak's event types so
/// BotEngine treats it exactly like Stimpak's client. Still experimental: only the test build turns
/// it on (<see cref="Enabled"/>), and it keeps a protocol log of every record it reads, with hex
/// for anything it can't decode, so a failed session says exactly where it stopped.
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

    /// <summary>Set to 1 to use this client instead of Stimpak's. The test build sets it.</summary>
    public const string EnableVariable = "INVIGORATION_NATIVE_SC2";

    /// <summary>StarCraft II's public "General" channel.</summary>
    public const ushort GeneralChannelId = 1033;

    private static readonly Dictionary<ushort, string> KnownPublicChannels = new()
    {
        [1033] = "General",
        [1034] = "Trade",
        [1035] = "Help",
    };

    public static bool Enabled => Environment.GetEnvironmentVariable(EnableVariable) == "1";

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
        var target = FindMember(name) ?? _self
            ?? throw new StimpakException("Can't whisper before a character is selected.", null!);
        Send(ChatCommands.ChatWhisper(new WhisperTarget.ToonName(name, target.Region, target.ProgramId, target.Realm), body), "whisper " + name);
        Emit(new WhisperReceived(name, body, true));
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
            await SignInAndChatAsync(channels, cancellationToken).ConfigureAwait(false);
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
        Emit(new StageChanged(Stage.WebAuthentication));
        var front = new FrontClient();
        var frontOpen = true;
        try
        {
            Trace($"Opening Front at {FrontClient.DefaultUsUri}");
            await front.ConnectAsync(new Uri(FrontClient.DefaultUsUri), cancellationToken).ConfigureAwait(false);
            await front.EstablishAsync(cancellationToken).ConfigureAwait(false);

            var saved = BattlenetCredentialProfileStore.LoadNativeCredential(_profileId, Program);
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

            Emit(new StageChanged(Stage.GameUtilities));
            var gameAccount = logon.GameAccountIds.FirstOrDefault()
                ?? throw new StimpakException("This Battle.net account has no StarCraft II game account.", null!);
            var sessionKey = logon.SessionKey ?? throw new StimpakException("Battle.net sent no session key.", null!);
            var handoff = await front.ProcessClientRequestAsync(gameAccount, sessionKey, cancellationToken).ConfigureAwait(false);
            Trace($"Handoff to Sunken at {handoff.Address}.");

            Emit(new StageChanged(Stage.NativeAuthentication));
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
        try
        {
            await ReadLoopAsync(channels, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            pinger.Dispose();
            await pinging.ConfigureAwait(ConfigureAwaitOptions.SuppressThrowing);
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
                    break;

                case NativeChatRecord.Message message:
                    HandleMessage(message.Value);
                    break;

                case NativeChatRecord.Whisper whisper:
                    Emit(new WhisperReceived(whisper.Value.PeerName, whisper.Value.Body, false));
                    break;
            }
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

    private async Task<byte[]> ChallengeAsync(Uri url, CancellationToken cancellationToken)
    {
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
