using System.Collections.ObjectModel;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;
using Stimpak;

namespace Invigoration.Core;

/// <summary>
/// StarCraft II/SC:Remastered/WC3:Reforged chat, backed by ncarrillo/superiority's Stimpak
/// native library — consumed as the <c>Stimpak</c> NuGet package (see
/// dotnet/src/StimpakPackage.props for the packaging story) rather than a hand-rolled port of
/// the native ("Sunken") protocol. Stimpak already implements the full protocol — including
/// startup records this project's own earlier hand-decoding effort in Invigoration.Sc2 never
/// finished reverse-engineering — behind a small, stable event stream, so this file is mostly
/// translation: Stimpak's <see cref="SC2Event"/>s in, the same shared
/// <see cref="BotEngine.HandleChatEvent"/> pipeline the classic BNCS and Chat/Telnet
/// connections use out, so roster tracking, clan tracking, trivia, and command dispatch all
/// work unmodified for an SC2 bot too (see BotEngine.Chat.cs's remarks on that pattern).
/// SC2/SC:R/WC3:R all connect identically — Stimpak's C# API has no per-game selector at all,
/// "supporting" one is purely a matter of its own protocol decoder correctly handling that
/// game's toon/presence data, not something this layer needs to branch on.
///
/// Toon selection happens inside Stimpak itself once connected. Unlike classic BNCS
/// (protocol-level single-channel), Stimpak supports being joined to multiple channels at once
/// — see <see cref="MaxJoinedSc2Channels"/> and the Sc2Channel* members below. Which channels
/// to restore on connect (empty means just the default "General") is Stimpak's own native
/// <see cref="StimpakConnectOptions.Channels"/> option, not something replayed by hand after
/// the fact — see ConnectSc2Async/PersistSc2ChannelList.
/// </summary>
public sealed partial class BotEngine
{
    /// <summary>Native/protocol hard cap on simultaneously joined SC2 channels (ncarrillo/superiority's MAX_JOINED_CHANNELS). Not exposed via Stimpak's FFI surface, so this has no compile-time tie to the native crate — keep in sync by hand if that constant ever changes.</summary>
    public const int MaxJoinedSc2Channels = 6;

    /// <summary>
    /// Set by the App layer before Connect is called on an SC2 bot — pops a Battle.net login
    /// dialog and returns the resulting web-auth credential. Stimpak's base package always
    /// surfaces <see cref="AuthenticationRequired"/> and leaves answering it entirely to the
    /// caller (an optional Stimpak.Auth package offers an in-process native WebView instead,
    /// but this app doesn't reference it — this handler is the only sign-in path).
    /// </summary>
    public Func<Uri, CancellationToken, Task<byte[]>>? Sc2ChallengeHandler { get; set; }

    /// <summary>Fired once a channel is joined, handing back Stimpak's own live per-channel roster (already correctly reconciled — see PeopleRegistry) so the UI can build a sub-tab in one step.</summary>
    public event Action<byte, ChatChannel, ObservableCollection<Person>>? Sc2ChannelJoined;

    /// <summary>Fired once a channel is actually left (confirmed by Stimpak, not merely requested) — the UI should close that channel's sub-tab.</summary>
    public event Action<byte>? Sc2ChannelLeft;

    /// <summary>A join attempt was rejected — either the server said no, or the local MaxJoinedSc2Channels cap was hit. Human-readable, ready to display.</summary>
    public event Action<string>? Sc2ChannelJoinRejected;

    /// <summary>Some other Stimpak-backed action failed (currently just LeaveSc2Channel) — human-readable, ready to display the same way as Sc2ChannelJoinRejected.</summary>
    public event Action<string>? Sc2ChannelActionFailed;

    /// <summary>The account's public-channel catalog, sent once per session — feeds the "join another channel" picker.</summary>
    public event Action<IReadOnlyList<ChatChannel>>? Sc2PublicChannelsReceived;

    private sealed record Sc2ChannelSession(byte ChannelIndex, ChatChannel Channel);

    private StimpakClient? _sc2Client;
    private CancellationTokenSource? _sc2ReceiveCts;
    private readonly Dictionary<byte, Sc2ChannelSession> _sc2Channels = new();

    /// <summary>Which channel an operator-typed send (as opposed to a reply to a specific incoming message) targets — kept in sync with whichever sub-tab the UI has focused.</summary>
    private byte? _sc2ActiveChannelIndex;

    /// <summary>Which channel the currently-running trivia round (if any) was started in — see HandleChatEvent's answer-matching gate in BotEngine.Bncs.cs.</summary>
    private byte? _sc2TriviaChannelIndex;

    private readonly Dictionary<string, FriendEntry> _sc2Friends = new();

    /// <summary>The account's public-channel catalog, cached from the last PublicChannelsReceived so a channel name (from the "join" bot-command) can be resolved to the id JoinPublic needs.</summary>
    private IReadOnlyList<ChatChannel> _sc2PublicChannelCatalog = [];

    /// <summary>Whether another SC2 channel can be joined right now — false once MaxJoinedSc2Channels is reached, or if this isn't a connected SC2 bot at all.</summary>
    public bool CanJoinAnotherSc2Channel =>
        Protocol.BncsProduct.IsStimpakBacked(Config.Product) && _sc2Client is not null && _sc2Channels.Count < MaxJoinedSc2Channels;

    public bool TryJoinSc2PublicChannel(ushort channelId)
    {
        if (!Protocol.BncsProduct.IsStimpakBacked(Config.Product))
        {
            return false;
        }

        // Checked before touching _sc2Client (and before it needs to be non-null) so the cap
        // is enforceable/testable independent of a live connection.
        if (_sc2Channels.Count >= MaxJoinedSc2Channels)
        {
            Sc2ChannelJoinRejected?.Invoke($"You can be in {MaxJoinedSc2Channels} channels at once. Close one to join another.");
            return false;
        }

        if (_sc2Client is not { } client)
        {
            return false;
        }

        try
        {
            client.JoinPublic(channelId);
            return true;
        }
        catch (StimpakException ex)
        {
            LogError($"Could not join channel: {ex.Message}");
            return false;
        }
    }

    public bool TryJoinSc2PrivateChannel(string name)
    {
        if (!Protocol.BncsProduct.IsStimpakBacked(Config.Product))
        {
            return false;
        }

        if (_sc2Channels.Count >= MaxJoinedSc2Channels)
        {
            Sc2ChannelJoinRejected?.Invoke($"You can be in {MaxJoinedSc2Channels} channels at once. Close one to join another.");
            return false;
        }

        if (_sc2Client is not { } client)
        {
            return false;
        }

        try
        {
            client.JoinPrivate(name);
            return true;
        }
        catch (StimpakException ex)
        {
            LogError($"Could not join channel: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Requests leaving a channel — removes it from local tracking (and fires Sc2ChannelLeft,
    /// closing the sub-tab) immediately after a successful native call, not by waiting for the
    /// async Left/Removed SC2Event. Confirmed by reading Stimpak's own Rust source
    /// (core/src/games/sc2/chat/session.rs, leave_channel): it forgets the channel from its
    /// *local* state synchronously, as part of the very same call that sends the wire packet —
    /// it does not wait for a server round-trip. The Left event this engine also listens for is
    /// driven by the *server* independently pushing a roster update that removes our own
    /// handle — that reliably happens if something else removes us (e.g. kicked), but there's
    /// no guarantee it follows a leave WE initiated, which left some channels' tabs (most
    /// visibly the always-auto-joined default one) stuck open indefinitely even though the
    /// leave had, in fact, already succeeded. Left's own handler is effectively just a fallback
    /// for the server-initiated case now — RemoveSc2Channel's own guard makes calling it twice
    /// for the same channel harmless either way.
    /// </summary>
    public void LeaveSc2Channel(byte channelIndex)
    {
        if (!Protocol.BncsProduct.IsStimpakBacked(Config.Product) || _sc2Client is not { } client)
        {
            return;
        }

        try
        {
            client.Leave(channelIndex);
        }
        catch (StimpakException ex)
        {
            // LogError alone reaches the flat log, which a SupportsMultiChannel bot hides
            // entirely — Sc2ChannelActionFailed also puts it where the operator can actually
            // see it (the active sub-tab's own chat log), same as Sc2ChannelJoinRejected does.
            LogError($"Could not leave channel: {ex.Message}");
            Sc2ChannelActionFailed?.Invoke($"Could not leave channel: {ex.Message}");
            return;
        }

        RemoveSc2Channel(channelIndex);
    }

    /// <summary>Drops a channel from local tracking and fires Sc2ChannelLeft — a no-op if it's already gone, so both LeaveSc2Channel's own immediate call and a later Left/SessionEnded event can safely call this for the same channel without double-firing.</summary>
    private void RemoveSc2Channel(byte channelIndex)
    {
        if (!_sc2Channels.Remove(channelIndex))
        {
            return;
        }

        if (_sc2ActiveChannelIndex == channelIndex)
        {
            _sc2ActiveChannelIndex = _sc2Channels.Keys.Cast<byte?>().FirstOrDefault();
        }

        PersistSc2ChannelList();
        Sc2ChannelLeft?.Invoke(channelIndex);
    }

    /// <summary>
    /// Resolves a typed channel name (from the "join" bot-command, or a remembered channel
    /// being replayed after reconnect) against the cached public-channel catalog — a match
    /// joins by id via JoinPublic, matching how the "+" flyout's own public-channel picker
    /// works; anything else is assumed to be a private channel name and joined via JoinPrivate
    /// directly, since that's the only other channel kind an operator can type a bare name for.
    /// </summary>
    private bool TryJoinSc2ChannelByName(string channelName)
    {
        if (_sc2PublicChannelCatalog.FirstOrDefault(c => string.Equals(c.Name, channelName, StringComparison.OrdinalIgnoreCase)) is PublicChannel match)
        {
            return TryJoinSc2PublicChannel(match.Id);
        }

        return TryJoinSc2PrivateChannel(channelName);
    }

    /// <summary>SC2/SC:R/WC3:R equivalent of the classic "join" bot-command — there's no server-side slash-command parser to hand a raw "/join" packet off to, so this resolves the name and calls the matching Stimpak API directly.</summary>
    private async Task HandleSc2JoinCommandAsync(string channelName, Func<string, Task> reply)
    {
        if (string.IsNullOrWhiteSpace(channelName))
        {
            await reply("Usage: join <channel name>").ConfigureAwait(false);
            return;
        }

        if (!TryJoinSc2ChannelByName(channelName))
        {
            await reply($"Could not join {channelName}.").ConfigureAwait(false);
        }
    }

    /// <summary>SC2/SC:R/WC3:R equivalent of a raw "/leave" — resolves a typed channel name against the channels this bot currently has joined and leaves the matching one (e.g. "leave Clan BNU" closes that sub-tab).</summary>
    private async Task HandleSc2LeaveCommandAsync(string channelName, Func<string, Task> reply)
    {
        if (string.IsNullOrWhiteSpace(channelName))
        {
            await reply("Usage: leave <channel name>").ConfigureAwait(false);
            return;
        }

        var match = _sc2Channels.Values.FirstOrDefault(s => string.Equals(s.Channel.Name, channelName, StringComparison.OrdinalIgnoreCase));
        if (match is null)
        {
            await reply($"Not in a channel named \"{channelName}\".").ConfigureAwait(false);
            return;
        }

        LeaveSc2Channel(match.ChannelIndex);
    }

    /// <summary>
    /// Stimpak's own ChatChannel (what a Joined event carries) as the ChannelTarget its
    /// Connect options want back (see ConnectSc2Async/StimpakConnectOptions.Channels) — null for
    /// a PartyChannel, which is only ever joined by accepting an invitation and so isn't
    /// something a later connect should try to restore.
    /// </summary>
    private static ChannelTarget? ToChannelTarget(ChatChannel channel) => channel switch
    {
        PublicChannel pub => ChannelTarget.Public(pub.Id),
        PrivateChannel priv => ChannelTarget.Private(priv.Name),
        GroupChannel group => ChannelTarget.Group(group.ClubId),
        _ => null,
    };

    /// <summary>
    /// Keeps Config.Sc2LastChannels in sync with the channels actually joined right now, so a
    /// later reconnect (or app restart) restores this same set — handed straight to Stimpak's
    /// own StimpakConnectOptions.Channels on the next ConnectSc2Async, not replayed by hand:
    /// Stimpak's native connect sequencing handles the always-auto-joined default channel
    /// itself, so there's no "already joined, don't double-join it" race to guard against here
    /// the way the old hand-rolled replay logic needed to.
    /// </summary>
    private void PersistSc2ChannelList()
    {
        var channels = _sc2Channels.Values
            .Select(s => ToChannelTarget(s.Channel))
            .OfType<ChannelTarget>()
            .ToList();

        if (channels.SequenceEqual(Config.Sc2LastChannels))
        {
            return;
        }

        Config.Sc2LastChannels = channels;
        ConfigPersistNeeded?.Invoke();
    }

    /// <summary>Called by the UI when the operator switches sub-tabs, so a typed message goes to the right channel.</summary>
    public void SetActiveSc2Channel(byte channelIndex)
    {
        if (_sc2Channels.ContainsKey(channelIndex))
        {
            _sc2ActiveChannelIndex = channelIndex;
        }
    }

    /// <summary>Where this bot's Stimpak session caches its signed-in credential.</summary>
    private string Sc2CredentialPath => BattlenetCredentialProfileStore.CredentialFilePath(EnsureBattlenetCredentialProfileId());

    /// <summary>
    /// Resolves this bot's assigned Battle.net credential profile, auto-creating
    /// one (named after the bot) if none is assigned yet or the assigned one
    /// was since deleted from Manage Battle.net Profiles — a connect should
    /// never fail purely for lack of somewhere to cache a session. Fires
    /// ConfigPersistNeeded so the newly-stamped id actually reaches bots.json
    /// rather than only living in memory until some unrelated save happens to
    /// occur (see ConfigPersistNeeded's remarks on BotEngine.cs).
    /// </summary>
    private string EnsureBattlenetCredentialProfileId()
    {
        if (!string.IsNullOrEmpty(Config.BattlenetCredentialProfileId) &&
            BattlenetCredentialProfileStore.Find(Config.BattlenetCredentialProfileId) is not null)
        {
            return Config.BattlenetCredentialProfileId;
        }

        var profile = BattlenetCredentialProfileStore.CreateAndSave(Config.DisplayName);
        Config.BattlenetCredentialProfileId = profile.Id;
        ConfigPersistNeeded?.Invoke();
        return profile.Id;
    }

    private Task ConnectSc2Async(CancellationToken cancellationToken)
    {
        StimpakNativeResolver.Register();
        LogInfo("Connecting to Battle.net (StarCraft II)...");
        Directory.CreateDirectory(Path.GetDirectoryName(Sc2CredentialPath)!);

        StimpakClient client;
        try
        {
            // ApplicationId is required but doesn't matter for us — CredentialPath overrides
            // the per-user cache location it would otherwise derive, since credential storage
            // is already fully owned by BattlenetCredentialProfileStore (see Sc2CredentialPath).
            client = new StimpakClient(new StimpakClientOptions("cc.bnet.invigoration") { CredentialPath = Sc2CredentialPath });
        }
        catch (Exception ex)
        {
            LogError($"StarCraft II connect failed: {ex.Message}");
            return Task.CompletedTask;
        }

        _sc2Client = client;
        _sc2Friends.Clear();
        _sc2Channels.Clear();
        _sc2ActiveChannelIndex = null;
        _sc2TriviaChannelIndex = null;
        _sc2PublicChannelCatalog = [];

        _sc2ReceiveCts = new CancellationTokenSource();
        var token = _sc2ReceiveCts.Token;
        _ = SafeSc2ConsumeLoopAsync(client, token);

        try
        {
            // Channels restores whatever this bot had joined last time — natively, on Stimpak's
            // own side, rather than the hand-rolled post-connect replay this used to be (which
            // had its own race with the always-auto-joined default channel — see the removed
            // MaybeRejoinRememberedSc2Channels for the history). An empty list here just means
            // "General", per StimpakConnectOptions.Channels' own doc comment.
            client.Connect(new StimpakConnectOptions
            {
                ForceInteractive = false,
                Channels = Config.Sc2LastChannels,
            });
        }
        catch (StimpakException ex)
        {
            LogError($"StarCraft II connect failed: {ex.Message}");
        }

        return Task.CompletedTask;
    }

    private async Task SafeSc2ConsumeLoopAsync(StimpakClient client, CancellationToken cancellationToken)
    {
        try
        {
            await foreach (var next in client.ReadEventsAsync(cancellationToken).ConfigureAwait(false))
            {
                await HandleSc2EventAsync(client, next).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException)
        {
            // Intentional disconnect — DisconnectAsync/DisposeAsync cancelled this loop.
        }
        catch (Exception ex)
        {
            LogError($"StarCraft II connection lost: {ex.Message}");
            BncsDisconnected?.Invoke(ex);
            MaybeScheduleAutoReconnect();
        }
    }

    private async Task HandleSc2EventAsync(StimpakClient client, SC2Event next)
    {
        client.People.Apply(next);
        switch (next)
        {
            case StageChanged { Stage: Stage.Connected }:
                // No manual join here — StimpakConnectOptions.Channels (passed at Connect time,
                // see ConnectSc2Async) already tells Stimpak natively which channels to restore,
                // General included if that list is empty.
                LogInfo("Connected — joining chat...");
                break;

            case StageChanged stage:
                LogDebug($"StarCraft II stage: {stage.Stage}");
                break;

            case AuthenticationRequired auth:
                await HandleSc2AuthenticationRequiredAsync(client, auth).ConfigureAwait(false);
                break;

            case AccountConnected connected:
                // Ties this profile to the real signed-in BattleTag so it's identifiable if the
                // user has more than one Battle.net account — see BattlenetCredentialProfile
                // .DisplayLabel. Config.BattlenetCredentialProfileId is already guaranteed set by
                // now (ConnectSc2Async resolves it via Sc2CredentialPath before Connect is called).
                BattlenetCredentialProfileStore.UpdateBattleTag(Config.BattlenetCredentialProfileId, connected.Account.BattleTag);

                // Account.Games is presumably which Blizzard products this account can actually
                // play (so a future check could refuse connecting a WC3:Reforged bot to an
                // account with no WC3:R license, matching what upstream's author described) — its
                // real string values haven't been observed yet against a live account, so this is
                // logged rather than acted on for now. Once a real value is seen here, wire an
                // actual gate instead of guessing at the format.
                LogDebug($"StarCraft II account: {connected.Account.BattleTag} — games: [{string.Join(", ", connected.Account.Games ?? [])}]");
                break;

            case Joined joined:
                var isFirstChannel = _sc2Channels.Count == 0;
                _sc2Channels[joined.ChannelIndex] = new Sc2ChannelSession(joined.ChannelIndex, joined.Channel);
                _sc2ActiveChannelIndex ??= joined.ChannelIndex;
                LogInfo($"Joined {joined.Channel.Name}.");
                Sc2ChannelJoined?.Invoke(joined.ChannelIndex, joined.Channel, client.People.Channel(joined.ChannelIndex));
                PersistSc2ChannelList();
                if (isFirstChannel)
                {
                    BncsConnected?.Invoke();
                }

                break;

            case JoinRejected rejected:
                var reason = rejected.Reason?.ToString() ?? "unknown";
                LogError($"Could not join StarCraft II chat (reason {reason}).");
                Sc2ChannelJoinRejected?.Invoke($"Could not join {rejected.Channel?.Name ?? "that channel"} (reason {reason}).");
                break;

            case Left left:
                // A server-initiated removal (e.g. kicked) — a self-initiated one already
                // removed this synchronously in LeaveSc2Channel, making this a harmless no-op
                // for that case (RemoveSc2Channel guards on it already being gone).
                RemoveSc2Channel(left.ChannelIndex);
                break;

            case MemberJoined member:
                await HandleChatEvent(new ChatEvent(ChatEventType.Join, member.User.Name, 0, 0, "", member.ChannelIndex)).ConfigureAwait(false);
                break;

            case MemberLeft member:
                await HandleChatEvent(new ChatEvent(ChatEventType.Leave, member.User.Name, 0, 0, "", member.ChannelIndex)).ConfigureAwait(false);
                break;

            case MessageReceived message:
                await HandleChatEvent(new ChatEvent(ChatEventType.Talk, message.Sender.Name, 0, 0, message.Body, message.ChannelIndex)).ConfigureAwait(false);
                break;

            case WhisperReceived { Outgoing: false } whisper:
                await HandleChatEvent(new ChatEvent(ChatEventType.Whisper, whisper.Peer, 0, 0, whisper.Body)).ConfigureAwait(false);
                break;

            // Confirmed via Stimpak's own Rust source (send_resolved_whisper,
            // core/src/games/sc2/chat/session.rs): a sent whisper pushes this Outgoing:true
            // event *synchronously*, in the same call that queues the wire send — not something
            // waiting on a server round-trip, so this is a reliable "your whisper actually went
            // out" confirmation, not just a best-effort echo. Without this case, a whisper reply
            // on SC2 was sent correctly but never showed up anywhere in this bot's own UI.
            case WhisperReceived { Outgoing: true } sent:
                await HandleChatEvent(new ChatEvent(ChatEventType.WhisperSent, sent.Peer, 0, 0, sent.Body)).ConfigureAwait(false);
                break;

            case WhisperFailed failed:
                LogError($"Whisper to {failed.Peer} failed: {failed.Reason}");
                break;

            case FriendsReceived friends:
                HandleSc2FriendsReceived(friends);
                break;

            case PublicChannelsReceived catalog:
                _sc2PublicChannelCatalog = catalog.Channels;
                Sc2PublicChannelsReceived?.Invoke(catalog.Channels);
                break;

            case CommandFailed failed:
                LogError($"StarCraft II command failed: {failed.Message}");
                break;

            case SessionFailed failed:
                LogError($"StarCraft II session failed: {failed.Message}");
                break;

            case SessionEnded:
                // Used to close every sub-tab silently and stop there — this bot then sat there
                // looking "connected" (no error, no BncsDisconnected) with a dead client
                // underneath: any further join/send just failed quietly. A real Battle.net
                // account can only run one live chat session at a time, so this reliably fires
                // whenever a second Stimpak-backed bot (SC2/SC:R/WC3:R) signs in with the *same*
                // shared Battle.net credential profile while this one is still connected — the
                // account's session gets handed to whichever client authenticated most recently,
                // and every other client sharing that profile is silently dropped. Surfacing it
                // properly (LogError + BncsDisconnected + the same auto-reconnect path a real
                // connection loss uses) at least makes that visible instead of silent — running
                // more than one Stimpak-backed bot on the *same* profile at the same time isn't
                // really supported by Battle.net itself, not something fixable purely client-side.
                foreach (var channelIndex in _sc2Channels.Keys.ToList())
                {
                    Sc2ChannelLeft?.Invoke(channelIndex);
                }

                _sc2Channels.Clear();
                _sc2ActiveChannelIndex = null;
                LogError("StarCraft II session ended — another bot may have signed in with the same Battle.net profile.");
                BncsDisconnected?.Invoke(null);
                MaybeScheduleAutoReconnect();
                break;

            // Not surfaced anywhere yet: roster snapshots (the UI binds Stimpak's own
            // PeopleRegistry.Channel(index) directly instead — see Sc2ChannelJoined),
            // group/party invitations.
            case RosterReceived or GroupInvitation or PartyInvitation or UnrecognisedEvent:
                break;
        }
    }

    private async Task HandleSc2AuthenticationRequiredAsync(StimpakClient client, AuthenticationRequired auth)
    {
        if (Sc2ChallengeHandler is null)
        {
            LogError("StarCraft II needs a sign-in, but no login window is available in this build.");
            return;
        }

        try
        {
            var url = new Uri(auth.Url);
            var credential = await Sc2ChallengeHandler(url, CancellationToken.None).ConfigureAwait(false);
            client.SubmitAuth(auth.AuthId, System.Text.Encoding.UTF8.GetString(credential));
        }
        catch (Exception ex)
        {
            LogError($"StarCraft II sign-in failed: {ex.Message}");
        }
    }

    private void HandleSc2FriendsReceived(FriendsReceived friends)
    {
        _sc2Friends.Clear();
        foreach (var friend in friends.Friends)
        {
            var (status, location) = friend.Presence switch
            {
                Presence.Away => (FriendStatus.Away, FriendLocation.InChat),
                Presence.Busy => (FriendStatus.DoNotDisturb, FriendLocation.InChat),
                Presence.InGame => (FriendStatus.None, FriendLocation.PublicGame),
                Presence.Available => (FriendStatus.None, FriendLocation.InChat),
                _ => (FriendStatus.None, FriendLocation.Offline),
            };
            _sc2Friends[friend.Name] = new FriendEntry(friend.Name, status, location, "sc2", "");
        }

        FriendsListUpdated?.Invoke(_sc2Friends.Values.ToList());
    }

    /// <summary>
    /// Unlike classic BNCS, Stimpak's chat has no server-side "/me" emote rendering — sent
    /// literally, it would just show up as the raw text "/me ..." in the channel. Approximates
    /// the same emote look with a plain *asterisk* line instead. A pure static method (rather
    /// than inlined in SendSc2Async) so the translation itself is directly unit-testable without
    /// a live Stimpak connection.
    /// </summary>
    public static string TranslateSc2EmoteText(string body) =>
        body.StartsWith("/me ", StringComparison.Ordinal) ? $"*{body[4..]}*" : body;

    /// <summary>
    /// The universal "/w username text" convention every whisper (an operator's reply-as-whisper,
    /// a clan rank's auto-whisper) is built with — see ReplyAsync/ApplyRankBehaviorsAsync. Parsed
    /// out here rather than at each caller so both keep working unchanged for classic BNCS/Chat-
    /// Telnet, where the server itself still parses a literal "/w" the normal way.
    /// </summary>
    public static bool TryParseSc2Whisper(string body, out string target, out string message)
    {
        target = "";
        message = "";
        if (!body.StartsWith("/w ", StringComparison.Ordinal))
        {
            return false;
        }

        var rest = body[3..];
        var spaceIndex = rest.IndexOf(' ');
        if (spaceIndex <= 0)
        {
            return false;
        }

        target = rest[..spaceIndex];
        message = rest[(spaceIndex + 1)..];
        return true;
    }

    /// <summary>Sends on <paramref name="channelOverride"/> if given (a reply to a specific incoming message), otherwise on the active/focused sub-tab's channel. A "/w username text" body is a whisper — not channel-scoped at all, so it bypasses the channel resolution entirely and goes through Stimpak's own SendWhisper instead of SendMessage.</summary>
    private Task SendSc2Async(string body, byte? channelOverride)
    {
        if (_sc2Client is not { } client)
        {
            return Task.CompletedTask;
        }

        if (TryParseSc2Whisper(body, out var target, out var message))
        {
            try
            {
                client.SendWhisper(target, message);
            }
            catch (StimpakException ex)
            {
                LogError($"Could not whisper {target}: {ex.Message}");
            }

            return Task.CompletedTask;
        }

        var channelIndex = channelOverride ?? _sc2ActiveChannelIndex;
        if (channelIndex is not { } idx)
        {
            return Task.CompletedTask;
        }

        body = TranslateSc2EmoteText(body);

        try
        {
            client.SendMessage(idx, body);
        }
        catch (StimpakException ex)
        {
            LogError($"Could not send: {ex.Message}");
        }

        return Task.CompletedTask;
    }

    private Task DisconnectSc2Async()
    {
        _sc2ReceiveCts?.Cancel();
        _sc2ReceiveCts = null;
        _sc2Channels.Clear();
        _sc2ActiveChannelIndex = null;
        _sc2TriviaChannelIndex = null;
        if (_sc2Client is { } client)
        {
            _sc2Client = null;
            client.Dispose();
        }

        return Task.CompletedTask;
    }
}
