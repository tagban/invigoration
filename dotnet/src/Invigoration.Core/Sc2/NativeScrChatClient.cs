using System.Text;
using System.Threading.Channels;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;
using Invigoration.Scr;
using Invigoration.Scr.LegacyChat;
using Stimpak;

namespace Invigoration.Core.Sc2;

/// <summary>
/// Invigoration's own StarCraft: Remastered connection (Invigoration.Scr), reporting in Stimpak's
/// event types so BotEngine treats it like any other SC2-family client. Stimpak itself only fakes
/// SC:R as SC2, so this is the only real one. Test build only for now, like
/// <see cref="NativeSc2ChatClient"/>, and it keeps the same kind of protocol log.
/// </summary>
/// <remarks>
/// SC:R is in one channel at a time, so there's one channel slot (<see cref="ChannelIndex"/>):
/// joining another channel leaves the current one, reported as Left then Joined. Channels are
/// reported by name (<see cref="PrivateChannel"/>) since SC:R's channel IDs change between
/// instances, and a name is what a reconnect needs to join it again. Members arrive as whole lists
/// in channel-list updates; joins and leaves are worked out by comparing them.
/// </remarks>
public sealed class NativeScrChatClient : ISc2ChatClient, IBattlenetFriendsClient
{
    /// <summary>The program code this client signs in as.</summary>
    public const string Program = ScrProtocol.ProgramCode;

    private const byte ChannelIndex = 0;

    private readonly string _profileId;
    private readonly uint _gateway;
    private readonly string _characterName;
    private readonly Action<string> _log;
    private readonly Channel<SC2Event> _events = Channel.CreateUnbounded<SC2Event>(new UnboundedChannelOptions { SingleReader = true });
    private readonly Channel<(Func<ScrConnection, CancellationToken, Task> Action, string What)> _outgoing =
        Channel.CreateUnbounded<(Func<ScrConnection, CancellationToken, Task>, string)>(new UnboundedChannelOptions { SingleReader = true });
    private readonly object _sync = new();
    private readonly Dictionary<ulong, TaskCompletionSource<string>> _pendingAuth = new();
    private readonly Dictionary<string, User> _members = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, uint> _handles = new(StringComparer.OrdinalIgnoreCase);
    private CancellationTokenSource? _runCts;
    private Task? _run;
    private ScrConnection? _connection;
    private ScrChannel? _channel;
    private ulong _nextAuthId;
    private uint _nextHandle;
    private StreamWriter? _protocolLog;
    private readonly Dictionary<ulong, ScrFriend> _friends = new();
    private readonly List<(uint AccountId, string Text)> _sentWhispers = [];
    private readonly Dictionary<ulong, ScrInvitation> _invitations = new();
    private readonly HashSet<string> _invited = new(StringComparer.OrdinalIgnoreCase);
    private Timer? _friendsTimer;
    private bool _everJoined;
    private readonly HashSet<string> _describedMembers = new(StringComparer.OrdinalIgnoreCase);
    private byte[]? _presented;
    private IDisposable? _signInTurn;
    private bool _rosterLoaded;
    private bool _disposed;

    public NativeScrChatClient(string profileId, uint gateway, string characterName, Action<string> log)
    {
        _profileId = profileId;
        _gateway = gateway;
        _characterName = characterName;
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
                throw new StimpakException("The StarCraft: Remastered connection is already running.", null!);
            }

            // Public targets are left over from Stimpak's SC2 stand-in: their IDs are SC2's.
            var home = options.Channels?.OfType<PrivateChannelTarget>().FirstOrDefault()?.Name ?? ScrConnectOptions.DefaultHomeChannel;
            _runCts = new CancellationTokenSource();
            _run = Task.Run(() => RunAsync(home, _runCts.Token));
        }
    }

    public void Disconnect()
    {
        lock (_sync)
        {
            _runCts?.Cancel();
        }
    }

    public void JoinPublic(ushort channelId)
    {
        var connection = RequireConnection("join a channel");
        var listed = connection.Channels.FirstOrDefault(c => c.Id == channelId)
            ?? throw new StimpakException($"Battle.net doesn't list channel {channelId}.", null!);
        JoinPrivate(listed.DisplayName);
    }

    /// <summary>Joins by name; a listed public channel by its ID. Waits for Battle.net to confirm before sending anything after it.</summary>
    public void JoinPrivate(string name) => Enqueue(
        async (connection, token) =>
        {
            try
            {
                await connection.JoinAsync(name, token).ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is TimeoutException or InvalidOperationException)
            {
                Trace($"Couldn't join {name}: {ex.Message}");
                Emit(new JoinRejected(new PrivateChannel(name), null));
            }
        },
        "join " + name);

    public void Leave(byte channelIndex) => Send(c => c.Chat.LeaveChannel(), "leave the channel");

    /// <summary>Sends to the channel, and shows it: Battle.net doesn't echo a sender's own lines.</summary>
    public void SendMessage(byte channelIndex, string body)
    {
        // Slash commands (/whois, /kick, /ban, /designate...) go to the server as commands; its
        // answer comes back as an info or error line.
        if (body.StartsWith('/') && body.Length > 1)
        {
            Send(c => c.Chat.SlashCommand(body), "send " + body.Split(' ')[0]);
            return;
        }

        Send(c => c.Chat.SendMessage(body), "send message");
        Emit(new MessageReceived(ChannelIndex, UserFor(_connection?.Toon?.Name ?? ""), body));
    }

    /// <summary>
    /// A Battle.net friend's BattleTag gets a Battle.net whisper, which reaches them whatever game
    /// they're in; any other name is a character, whispered in classic chat.
    /// </summary>
    public void SendWhisper(string name, string body)
    {
        if (FriendAccount(name) is { } friend)
        {
            lock (_sync)
            {
                _sentWhispers.Add((friend.AccountId, body));
                if (_sentWhispers.Count > 20)
                {
                    _sentWhispers.RemoveAt(0);
                }
            }

            Enqueue(
                (c, token) => c.RequestAsync(ScrWhispers.Service, ScrWhispers.SendWhisperMethod, ScrWhispers.SendRequest(friend.AccountId, body), token),
                "whisper " + friend.BattleTag);
            Emit(new WhisperReceived(friend.BattleTag, body, true));
            return;
        }

        Send(c => c.Chat.Whisper(name, body), "whisper " + name);
        Emit(new WhisperReceived(name, body, true));
    }

    /// <summary>The Battle.net friend with this BattleTag, if there is one.</summary>
    private (uint AccountId, string BattleTag)? FriendAccount(string name)
    {
        if (!name.Contains('#'))
        {
            return null;
        }

        lock (_sync)
        {
            return _friends.Values.FirstOrDefault(f => f.BattleTag.Equals(name.Trim(), StringComparison.OrdinalIgnoreCase)) is { } friend
                ? ((uint)friend.AccountId, friend.BattleTag)
                : null;
        }
    }

    /// <summary>
    /// A Battle.net whisper to or from us. Battle.net echoes the ones we send, including from
    /// StarCraft: Remastered itself or the Battle.net app; ours are already shown.
    /// </summary>
    private void HandleBattlenetWhisper(ScrWhisper whisper)
    {
        string who;
        lock (_sync)
        {
            if (whisper.Outgoing && _sentWhispers.Remove((whisper.AccountId, whisper.Text)))
            {
                return;
            }

            who = _friends.TryGetValue(whisper.AccountId, out var friend) ? friend.BattleTag
                : BattlenetFriendDirectory.Find(_profileId, whisper.AccountId)?.BattleTag
                ?? $"Battle.net account {whisper.AccountId}";
        }

        Emit(new WhisperReceived(who, whisper.Text, whisper.Outgoing));
    }

    public void AddFriend(string battleTag)
    {
        var tag = battleTag.Trim();
        lock (_sync)
        {
            _invited.Add(tag);
        }

        Enqueue(
            async (c, token) =>
            {
                var reply = await c.RequestAsync(ScrFriends.Service, ScrFriends.SendInvitationMethod, ScrFriends.SendInvitationRequest(tag), token).ConfigureAwait(false);
                Trace($"Friend request to {tag} sent; reply {Convert.ToHexString(reply)}.");
                Emit(new NativeChatEvent(new ChatEvent(ChatEventType.Info, "", 0, 0, $"Friend request sent to {tag}.")));
            },
            "send a friend request to " + tag);
    }

    public void RemoveFriend(string battleTag)
    {
        var friend = FriendAccount(battleTag)
            ?? throw new StimpakException($"{battleTag} isn't on this account's Battle.net friends list.", null!);
        Enqueue(
            async (c, token) =>
            {
                var reply = await c.RequestAsync(ScrFriends.Service, ScrFriends.RemoveFriendMethod, ScrFriends.RemoveFriendRequest(friend.AccountId), token).ConfigureAwait(false);
                Trace($"Removed {friend.BattleTag} (account {friend.AccountId}); reply {Convert.ToHexString(reply)}.");
                Emit(new NativeChatEvent(new ChatEvent(ChatEventType.Info, "", 0, 0, $"Removed {friend.BattleTag} from your Battle.net friends.")));
            },
            "remove " + friend.BattleTag + " from friends");
    }

    public void AnswerInvitation(ulong invitationId, bool accept)
    {
        string who;
        lock (_sync)
        {
            who = _invitations.TryGetValue(invitationId, out var invitation) ? invitation.BattleTag : $"request {invitationId}";
        }

        Enqueue(
            async (c, token) =>
            {
                var method = accept ? ScrFriends.AcceptInvitationMethod : ScrFriends.DeclineInvitationMethod;
                var reply = await c.RequestAsync(ScrFriends.Service, method, ScrFriends.AnswerInvitationRequest(invitationId), token).ConfigureAwait(false);
                Trace($"{(accept ? "Accepted" : "Declined")} {who}'s friend request; reply {Convert.ToHexString(reply)}.");
            },
            (accept ? "accept " : "decline ") + who + "'s friend request");
    }

    /// <summary>
    /// A friend request came or went. Battle.net's update doesn't say which way a request goes, so
    /// ones to BattleTags this bot invited count as sent.
    /// </summary>
    private void HandleInvitation(ScrInvitation invitation, bool removed)
    {
        Trace($"Friend request {invitation.Id} {(removed ? "gone" : "from/to")} {invitation.BattleTag}.");
        bool isNew;
        bool sent;
        List<FriendInvitation> all;
        lock (_sync)
        {
            sent = _invited.Contains(invitation.BattleTag);
            isNew = !removed && _invitations.TryAdd(invitation.Id, invitation);
            if (removed)
            {
                _invitations.Remove(invitation.Id);
            }

            all = _invitations.Values
                .Select(i => new FriendInvitation(i.Id, i.BattleTag, _invited.Contains(i.BattleTag)))
                .OrderBy(i => i.BattleTag, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        if (isNew && !sent && _everJoined)
        {
            Emit(new NativeChatEvent(new ChatEvent(ChatEventType.Info, "", 0, 0,
                $"{invitation.BattleTag} sent you a Battle.net friend request. Accept or decline it on the Friends tab.")));
        }

        Emit(new NativeInvitationsEvent(all));
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
        });
    }

    private async Task RunAsync(string homeChannel, CancellationToken cancellationToken)
    {
        OpenProtocolLog();
        using var session = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        Task? sending = null;
        try
        {
            Emit(new StageChanged(Stage.WebAuthentication));
            Trace($"Gateway {ScrGateways.NameOf(_gateway)}, character {(_characterName.Length > 0 ? _characterName : "(first there)")}, home channel {homeChannel}.");
            _connection = await ConnectWithRetryAsync(homeChannel, cancellationToken).ConfigureAwait(false);

            // After startup, which Battle.net wants done within seconds: other games' sign-ins, then
            // let any other game on this profile waiting to sign in have its turn.
            await NativeSignIns.IssueMissingAsync(_profileId, Program, _connection.GenerateWebCredentialsAsync, Trace, cancellationToken).ConfigureAwait(false);
            EndSignInTurn();

            Emit(new AccountConnected(new AccountSummary(null, _connection.BattleTag ?? "", null, [])));
            Emit(new PublicChannelsReceived(_connection.Channels
                .Where(c => c.Id <= ushort.MaxValue)
                .OrderBy(c => c.Id)
                .Select(c => (ChatChannel)new PublicChannel((ushort)c.Id, c.DisplayName))
                .ToList()));

            sending = SendLoopAsync(_connection, session.Token);
            await foreach (var chatEvent in _connection.Events.ReadAllAsync(cancellationToken).ConfigureAwait(false))
            {
                Handle(chatEvent);
            }

            throw new IOException("Battle.net closed the StarCraft: Remastered connection.");
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            Trace("Disconnected on request.");
        }
        catch (Exception ex)
        {
            Trace($"Session failed: {ex}");
            Emit(new SessionFailed(ex.Message));
        }
        finally
        {
            EndSignInTurn();
            await session.CancelAsync().ConfigureAwait(false);
            if (sending is not null)
            {
                await sending.ConfigureAwait(ConfigureAwaitOptions.SuppressThrowing);
            }

            if (_connection is { } connection)
            {
                _connection = null;
                await connection.DisposeAsync().ConfigureAwait(false);
            }

            lock (_sync)
            {
                foreach (var pending in _pendingAuth.Values)
                {
                    pending.TrySetCanceled();
                }

                _pendingAuth.Clear();
                _members.Clear();
                _channel = null;
                _friends.Clear();
                _invitations.Clear();
                _friendsTimer?.Dispose();
                _friendsTimer = null;
                BattlenetFriendDirectory.ForgetActivity(_profileId);
            }

            Emit(new StageChanged(Stage.Disconnected));
            _protocolLog?.Dispose();
            _protocolLog = null;
        }
    }

    /// <summary>
    /// Signs in and starts the game session, trying again if Battle.net drops the connection
    /// during startup. It does that while it still considers the account's last SC:R session live
    /// (the character list never comes) for 20-50 seconds after it ends, so the retries come every
    /// 10 seconds, six at most.
    /// </summary>
    private async Task<ScrConnection> ConnectWithRetryAsync(string homeChannel, CancellationToken cancellationToken)
    {
        // Battle.net lets go of the last session 20-50 seconds after it ends; retry every 10 until then.
        TimeSpan[] waits = [.. Enumerable.Repeat(TimeSpan.FromSeconds(10), 6)];
        for (var attempt = 0; ; attempt++)
        {
            try
            {
                // Reloaded each time: the attempt before may have saved a newer one.
                var saved = _presented = BattlenetCredentialProfileStore.LoadNativeCredential(_profileId, Program);
                return await ScrConnection.ConnectAsync(
                    saved,
                    new ScrConnectOptions(_gateway, _characterName, homeChannel),
                    ChallengeAsync,
                    Trace,
                    cancellationToken,
                    fresh => BattlenetCredentialProfileStore.SaveNativeCredential(_profileId, Program, fresh)).ConfigureAwait(false);
            }
            catch (SavedSignInArrivedException arrived)
            {
                // Not a failure: another game on the profile signed in for us meanwhile.
                Trace(arrived.Message);
                attempt--;
            }
            catch (Exception ex) when (attempt < waits.Length && IsStartupDrop(ex) && !cancellationToken.IsCancellationRequested)
            {
                var wait = waits[attempt];
                Trace($"Battle.net dropped the connection during startup ({ex.Message}), likely still closing this account's last StarCraft: Remastered session. Trying again in {wait.TotalSeconds:0} seconds.");
                Emit(new NativeChatEvent(new ChatEvent(ChatEventType.Info, "", 0, 0,
                    $"Battle.net is still closing this account's last StarCraft: Remastered session. Trying again in {wait.TotalSeconds:0} seconds.")));
                await Task.Delay(wait, cancellationToken).ConfigureAwait(false);
            }
        }
    }

    /// <summary>The classic connection closed or went quiet mid-startup, as opposed to a refused sign-in or a missing character.</summary>
    private static bool IsStartupDrop(Exception ex) =>
        ex is TimeoutException || (ex is InvalidOperationException && ex is not ScrCharacterException && ex.Message.Contains("classic connection closed", StringComparison.Ordinal));

    private void Handle(ScrChatEvent chatEvent)
    {
        switch (chatEvent)
        {
            case ScrChatEvent.ChannelEntered entered:
                Enter(entered.Channel);
                break;

            case ScrChatEvent.ChannelListChanged list:
                foreach (var change in list.Changes)
                {
                    if (!change.IsRemoval && _channel is { } current && IsSameChannel(current, change.Channel))
                    {
                        UpdateMembers(change.Channel.Members);
                    }
                }

                break;

            case ScrChatEvent.ChannelLeft left when _channel is { } current && (left.ChannelId == current.Id || left.ChannelId == 0):
                _channel = null;
                Emit(new Left(ChannelIndex, null));
                break;

            case ScrChatEvent.Message { Value: var message }:
                HandleMessage(message);
                break;

            case ScrChatEvent.Unhandled { Service: ScrFriends.Service, Method: ScrFriends.FriendUpdatedMethod } friendUpdate:
                if (ScrFriends.DecodeUpdate(friendUpdate.Body) is var (friend, removed))
                {
                    if (friend.Online)
                    {
                        _protocolLog?.WriteLine($"{DateTime.Now:HH:mm:ss.fff}    friend {friend.BattleTag}: program {friend.Program}, detail \"{friend.Detail}\", flags {friend.Away}/{friend.Busy}");
                    }

                    lock (_sync)
                    {
                        if (removed)
                        {
                            _friends.Remove(friend.AccountId);
                        }
                        else
                        {
                            _friends[friend.AccountId] = friend;
                        }

                        // Battle.net sends the whole list one friend at a time at sign-in: report it once it settles.
                        _friendsTimer ??= new Timer(_ => EmitFriends());
                        _friendsTimer.Change(TimeSpan.FromMilliseconds(300), Timeout.InfiniteTimeSpan);
                    }
                }

                break;

            case ScrChatEvent.Unhandled { Service: ScrFriends.Service, Method: ScrFriends.InvitationUpdatedMethod } invitationUpdate:
                if (ScrFriends.DecodeInvitation(invitationUpdate.Body) is var (invitation, gone))
                {
                    HandleInvitation(invitation, gone);
                }

                break;

            case ScrChatEvent.Unhandled { Service: ScrWhispers.Service } aurora when ScrWhispers.Decode(aurora.Method, aurora.Body) is { } whisper:
                HandleBattlenetWhisper(whisper);
                break;

            case ScrChatEvent.Unhandled unhandled:
                _protocolLog?.WriteLine($"{DateTime.Now:HH:mm:ss.fff}    body {unhandled.Service:X8}/{unhandled.Method:X8}: {Convert.ToHexString(unhandled.Body)}");
                break;
        }
    }

    private void EmitFriends()
    {
        List<FriendEntry> friends;
        lock (_sync)
        {
            BattlenetFriendDirectory.Update(_profileId, _friends.Values.Select(f =>
                ((uint)f.AccountId, f.BattleTag, f.RealName, new BattlenetFriendDirectory.Activity(f.Online, f.Program, f.Detail))));
            friends = _friends.Values
                .OrderByDescending(f => f.Online).ThenBy(f => f.BattleTag, StringComparer.OrdinalIgnoreCase)
                .Select(ToFriendEntry)
                .ToList();
        }

        Trace($"Friends list: {friends.Count} ({friends.Count(f => f.Location != FriendLocation.Offline)} online).");
        Emit(new NativeFriendsEvent(friends));
    }

    /// <summary>A Battle.net friend as the Battle.net app shows them: BattleTag, real name if shared, and the game they're in.</summary>
    private static FriendEntry ToFriendEntry(ScrFriend friend)
    {
        var status = friend.Busy ? FriendStatus.DoNotDisturb : friend.Away ? FriendStatus.Away : FriendStatus.None;
        if (!friend.Online)
        {
            return new FriendEntry(friend.BattleTag, status, FriendLocation.Offline, "", "", friend.RealName);
        }

        var (icon, game) = BattlenetPrograms.Describe(friend.Program);
        var doing = friend.Detail.Length > 0 ? $"{game}: {friend.Detail}" : game;
        return new FriendEntry(friend.BattleTag, status, FriendLocation.NotInChat, icon, doing, friend.RealName);
    }

    private void Enter(ScrChannel channel)
    {
        var first = !_everJoined;
        if (_channel is not null)
        {
            Emit(new Left(ChannelIndex, null));
        }

        lock (_sync)
        {
            _channel = channel;
            _members.Clear();
            _rosterLoaded = false;
        }

        var name = channel.DisplayName.Length > 0 ? channel.DisplayName : channel.InternalName;
        Emit(new Joined(ChannelIndex, new PrivateChannel(name), 0));

        // Usually empty: the member list follows in a channel-list update, often a second later but
        // sometimes not for 20 seconds or more. If it hasn't come in two, ask for the channel list
        // again, which may bring it sooner.
        if (channel.Members.Count > 0)
        {
            UpdateMembers(channel.Members);
        }
        else
        {
            _ = Task.Delay(TimeSpan.FromSeconds(2)).ContinueWith(_ =>
            {
                bool waiting;
                lock (_sync)
                {
                    waiting = ReferenceEquals(_channel, channel) && !_rosterLoaded;
                }

                if (waiting && _connection is not null)
                {
                    Trace($"No member list for {name} yet; asking for the channel list again.");
                    Send(c => c.Chat.ListChannels(), "ask for the channel list");
                }
            }, TaskScheduler.Default);
        }
        if (first)
        {
            _everJoined = true;
            Emit(new StageChanged(Stage.Connected));
        }
    }

    /// <summary>
    /// Channel-list updates carry the whole member list. The first one after entering a channel is
    /// who was already there, reported as the roster; after that the difference is who came and went.
    /// </summary>
    private void UpdateMembers(IReadOnlyList<ScrMember> members)
    {
        foreach (var member in members)
        {
            // SC:R chat only meets Warcraft II, StarCraft, Diablo (in some channels) and Diablo II,
            // and every one but Diablo II comes with its code; so no code means Diablo II. Whether
            // it's Lord of Destruction isn't said, so it's the plain Diablo II icon.
            NativeMemberProducts.Set(member.Name, member.Attributes.TryGetValue("program_id", out var program) ? program : "D2DV");
            if (member.Attributes.TryGetValue("battle_tag", out var battleTag))
            {
                NativeMemberProducts.SetBattleTag(member.Name, battleTag);
            }

            // What each member carries, once per name, to learn what other games send (icons).
            if (_describedMembers.Add(member.Name))
            {
                _protocolLog?.WriteLine($"{DateTime.Now:HH:mm:ss.fff}    member {member.Name}: flags {member.Flags}; "
                    + string.Join(", ", member.Attributes.Where(a => a.Key != "battle_tag").Select(a => $"{a.Key}={a.Value}")));
            }
        }

        var now = members.Select(m => m.Name).Where(n => n.Length > 0).ToHashSet(StringComparer.OrdinalIgnoreCase);
        List<User> joined = [], left = [];
        bool initial;
        lock (_sync)
        {
            initial = !_rosterLoaded;
            _rosterLoaded = true;
            foreach (var name in now.Where(n => !_members.ContainsKey(n)))
            {
                var user = UserFor(name);
                _members[name] = user;
                joined.Add(user);
            }

            foreach (var name in _members.Keys.Where(n => !now.Contains(n)).ToList())
            {
                left.Add(_members[name]);
                _members.Remove(name);
            }
        }

        if (initial)
        {
            Emit(new RosterReceived(ChannelIndex, true, joined));
            return;
        }

        foreach (var user in joined)
        {
            Emit(new MemberJoined(ChannelIndex, user));
        }

        foreach (var user in left)
        {
            Emit(new MemberLeft(ChannelIndex, user));
        }
    }

    private void HandleMessage(ScrMessage message)
    {
        var sender = message.Sender ?? "";
        switch (message.Kind)
        {
            case ScrMessageKind.Channel:
                Emit(new MessageReceived(ChannelIndex, UserFor(sender), message.Text));
                break;
            case ScrMessageKind.Whisper:
                Emit(new WhisperReceived(sender, message.Text, false));
                break;
            case ScrMessageKind.Emote:
                Emit(new NativeChatEvent(new ChatEvent(ChatEventType.Emote, sender, 0, 0, message.Text, ChannelIndex)));
                break;
            case ScrMessageKind.Broadcast:
                Emit(new NativeChatEvent(new ChatEvent(ChatEventType.Broadcast, sender.Length > 0 ? sender : "Battle.net", 0, 0, message.Text, ChannelIndex)));
                break;
            case ScrMessageKind.Information:
                Emit(new NativeChatEvent(new ChatEvent(ChatEventType.Info, "", 0, 0, message.Text, ChannelIndex)));
                break;
            case ScrMessageKind.Error:
                Emit(new NativeChatEvent(new ChatEvent(ChatEventType.Error, "", 0, 0, message.Text, ChannelIndex)));
                break;
        }
    }

    private static bool IsSameChannel(ScrChannel current, ScrChannel other) =>
        other.Id == current.Id
        || (current.DisplayName.Length > 0 && other.DisplayName.Equals(current.DisplayName, StringComparison.OrdinalIgnoreCase));

    /// <summary>SC:R has no member handles; each name gets a stable one for this client's lifetime.</summary>
    private User UserFor(string name)
    {
        lock (_sync)
        {
            if (!_handles.TryGetValue(name, out var handle))
            {
                handle = ++_nextHandle;
                _handles[name] = handle;
            }

            return new User(handle, null, name, "", Presence.Available);
        }
    }

    private ScrConnection RequireConnection(string what) =>
        _connection ?? throw new StimpakException($"Can't {what}: not connected to StarCraft: Remastered chat.", null!);

    private void Send(Func<ScrConnection, byte[]> build, string what) =>
        Enqueue((connection, token) => connection.SendAsync(build(connection), token), what);

    /// <summary>Queues an action for the send loop, which runs them one at a time, in order.</summary>
    private void Enqueue(Func<ScrConnection, CancellationToken, Task> action, string what)
    {
        RequireConnection(what);
        _outgoing.Writer.TryWrite((action, what));
    }

    private async Task SendLoopAsync(ScrConnection connection, CancellationToken cancellationToken)
    {
        try
        {
            await foreach (var (action, what) in _outgoing.Reader.ReadAllAsync(cancellationToken).ConfigureAwait(false))
            {
                try
                {
                    _protocolLog?.WriteLine($"{DateTime.Now:HH:mm:ss.fff} -> {what}");
                    await action(connection, cancellationToken).ConfigureAwait(false);
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    Trace($"Couldn't {what}: {ex.Message}");
                    Emit(new CommandFailed($"Couldn't {what}: {ex.Message}"));
                }
            }
        }
        catch (OperationCanceledException)
        {
            // Session over.
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

        Trace($"Battle.net asked for a web sign-in at {url.Scheme}://{url.Host}{url.AbsolutePath}.");
        Emit(new AuthenticationRequired(authId, url.ToString(), false));
        using var registration = cancellationToken.Register(() => answer.TrySetCanceled(cancellationToken));
        var token = await answer.Task.ConfigureAwait(false);
        return Encoding.UTF8.GetBytes(token);
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
            var path = Path.Combine(directory, $"scr-native-{DateTime.Now:yyyyMMdd-HHmmss}.log");
            _protocolLog = new StreamWriter(path, append: false, Encoding.UTF8) { AutoFlush = true };
            _log($"Native StarCraft: Remastered protocol log: {path}");
        }
        catch (Exception ex)
        {
            _log($"Couldn't open the native StarCraft: Remastered protocol log: {ex.Message}");
        }
    }

    /// <summary>A line for both the bot's debug log and the protocol log. Never given credentials or keys.</summary>
    private void Trace(string message)
    {
        _log($"[native SC:R] {message}");
        _protocolLog?.WriteLine($"{DateTime.Now:HH:mm:ss.fff} {message}");
    }
}
