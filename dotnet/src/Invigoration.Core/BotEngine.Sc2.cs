using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;
using Invigoration.Core.Sc2;
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
    /// <summary>
    /// How many chat channels an SC2, SC:R or WC3:R bot can have open at once. Six, matching Stimpak's
    /// limit (the protocol itself has seven slots, a 3-bit index). The games may stop at five; this
    /// is to be confirmed against a live account before changing it.
    /// </summary>
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

    private ISc2ChatClient? _sc2Client;

    /// <summary>Cancelled when the current client is retired, closing a login window it has open. Only the sign-in uses it; the client's events are read until it's disposed.</summary>
    private CancellationTokenSource? _sc2ReceiveCts;

    /// <summary>The Battle.net profile the current client saves its sign-in to. Fixed when the client is made, so a bot moved to another profile needs a new client.</summary>
    private string? _sc2ClientProfileId;

    /// <summary>This bot's hold on its Battle.net profile's sign-in, from a connect until that attempt or session is over. See BattlenetSignInLease.</summary>
    private BattlenetSignInLease? _sc2Lease;

    /// <summary>Retired clients still finishing an attempt, each with what completes once Stimpak says it has stopped.</summary>
    private readonly ConcurrentDictionary<ISc2ChatClient, TaskCompletionSource> _sc2WindingDown = new();

    /// <summary>
    /// Longest a retired client gets to finish the attempt it was making before it's released
    /// anyway. Stimpak only takes a Disconnect once it's in chat, and each stage before that can take
    /// about 30 seconds.
    /// </summary>
    private static readonly TimeSpan Sc2WindDownTimeout = TimeSpan.FromMinutes(2);

    /// <summary>How old the saved sign-in was when this attempt started, or null if there was none. Said if Battle.net turns it down, which is how its real lifetime shows up in the log.</summary>
    private TimeSpan? _sc2SavedSignInAge;

    /// <summary>Sessions lost in a row soon after getting into chat. See NoteSc2SessionLost.</summary>
    private int _sc2QuickSessionLosses;

    private const int MaxSc2QuickSessionLosses = 3;

    /// <summary>
    /// In chat: Stimpak said Connected (chat set up, first channel joined) and hasn't since said
    /// Disconnected. SC2's equivalent of classic Battle.net's logged-on flag — what ends an
    /// auto-reconnect, and what tells a Disconnected that a live session was lost rather than an
    /// attempt failing.
    /// </summary>
    private volatile bool _sc2InChat;

    /// <summary>The current client's worker has gone (SessionEnded): it can't connect again, so the next connect needs a new one.</summary>
    private bool _sc2ClientEnded;
    private readonly Dictionary<byte, Sc2ChannelSession> _sc2Channels = new();

    /// <summary>Which channel an operator-typed send (as opposed to a reply to a specific incoming message) targets — kept in sync with whichever sub-tab the UI has focused.</summary>
    private byte? _sc2ActiveChannelIndex;

    /// <summary>Which channel the currently-running trivia round (if any) was started in — see HandleChatEvent's answer-matching gate in BotEngine.Bncs.cs.</summary>
    private byte? _sc2TriviaChannelIndex;

    private readonly Dictionary<string, FriendEntry> _sc2Friends = new();

    /// <summary>The StarCraft II clans and groups the bot's character belongs to (native SC2 only).</summary>
    public event Action<IReadOnlyList<BattlenetClub>>? ClubsUpdated;

    /// <summary>Opens a StarCraft II clan's or group's chat as a channel tab.</summary>
    public void JoinClubChat(uint clubId)
    {
        if (_sc2Client is NativeSc2ChatClient native)
        {
            native.JoinClubChat(clubId);
        }
        else
        {
            LogError("Clan chat needs Invigoration's own StarCraft II connection.");
        }
    }

    /// <summary>Pending Battle.net friend requests, from a native client that reports them (SC:R).</summary>
    public event Action<IReadOnlyList<FriendInvitation>>? FriendInvitationsUpdated;

    /// <summary>Whether this bot can change its Battle.net friends list: a connected SC:R bot, so far.</summary>
    public bool CanManageBattlenetFriends => _sc2Client is IBattlenetFriendsClient;

    /// <summary>Sends a Battle.net friend request to <paramref name="battleTag"/>. False if it isn't a BattleTag or can't be sent.</summary>
    public bool AddBattlenetFriend(string battleTag)
    {
        var tag = battleTag.Trim();
        var hash = tag.IndexOf('#');
        if (hash <= 0 || hash == tag.Length - 1 || tag.Contains(' '))
        {
            LogError("Friend requests go to a BattleTag, like Name#1234.");
            return false;
        }

        return WithFriendsClient(c => c.AddFriend(tag), $"send {tag} a friend request");
    }

    public bool RemoveBattlenetFriend(string battleTag) => WithFriendsClient(c => c.RemoveFriend(battleTag), $"remove {battleTag}");

    public bool AnswerFriendInvitation(ulong invitationId, bool accept) =>
        WithFriendsClient(c => c.AnswerInvitation(invitationId, accept), accept ? "accept the friend request" : "decline the friend request");

    private bool WithFriendsClient(Action<IBattlenetFriendsClient> action, string what)
    {
        if (_sc2Client is not IBattlenetFriendsClient client)
        {
            LogError($"Can't {what}: only a connected StarCraft: Remastered bot can change Battle.net friends so far.");
            return false;
        }

        try
        {
            action(client);
            return true;
        }
        catch (StimpakException ex)
        {
            LogError($"Couldn't {what}: {ex.Message}");
            return false;
        }
    }

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

        // On the Battle.net login, per game, so any bot signing in with it gets them back.
        if (!string.IsNullOrEmpty(Config.BattlenetCredentialProfileId))
        {
            BattlenetCredentialProfileStore.SetLastChannels(Config.BattlenetCredentialProfileId, Config.Product, channels);
        }

        if (channels.SequenceEqual(Config.Sc2LastChannels))
        {
            return;
        }

        Config.Sc2LastChannels = channels;
        ConfigPersistNeeded?.Invoke();
    }

    /// <summary>
    /// The channels to restore on connect: this login's for this game, or, for a login that hasn't
    /// saved any yet, the bot's own last list from before channels were kept per login.
    /// </summary>
    private IReadOnlyList<ChannelTarget> LastSc2Channels(string profileId) =>
        BattlenetCredentialProfileStore.LastChannels(profileId, Config.Product) ?? Config.Sc2LastChannels;

    /// <summary>Called by the UI when the operator switches sub-tabs, so a typed message goes to the right channel.</summary>
    public void SetActiveSc2Channel(byte channelIndex)
    {
        if (_sc2Channels.ContainsKey(channelIndex))
        {
            _sc2ActiveChannelIndex = channelIndex;
        }
    }

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

    /// <remarks>
    /// One client serves every attempt while it can: after any attempt or session ends (Stimpak's
    /// Disconnected stage) the same client connects again, which is what Stimpak intends. A new client
    /// is only made the first time, after SessionEnded (its worker gone), when the bot has been moved
    /// to another Battle.net profile, or when the current one is somehow still busy. A busy one is
    /// left to finish before it lets go of the sign-in (see RetireSc2Client).
    /// </remarks>
    private async Task ConnectSc2Async(CancellationToken cancellationToken)
    {
        StimpakNativeResolver.Register();
        LogInfo("Connecting to Battle.net (StarCraft II)...");
        var profileId = EnsureBattlenetCredentialProfileId();

        if (_sc2Client is not null &&
            (_sc2ClientEnded || Volatile.Read(ref _sc2LiveClient) is not null || _sc2ClientProfileId != profileId))
        {
            RetireSc2Client();
        }

        CloseSc2Channels();
        _sc2InChat = false;
        _sc2Friends.Clear();
        FriendInvitationsUpdated?.Invoke([]);
        ClubsUpdated?.Invoke([]);
        _sc2ActiveChannelIndex = null;
        _sc2TriviaChannelIndex = null;
        _sc2PublicChannelCatalog = [];

        // Nothing is live by now, so a hold still here has outlived its attempt.
        Interlocked.Exchange(ref _sc2Lease, null)?.Dispose();
        if (await AcquireSc2SignInAsync(profileId, cancellationToken).ConfigureAwait(false) is not { } lease)
        {
            return;
        }

        if (_sc2Client is not { } client)
        {
            try
            {
                client = CreateSc2Client(profileId);
            }
            catch (Exception ex)
            {
                lease.Dispose();
                LogError($"StarCraft II connect failed: {ex.Message}");
                return;
            }

            _sc2Client = client;
            _sc2ClientProfileId = profileId;
            _sc2ClientEnded = false;
            _sc2ReceiveCts = new CancellationTokenSource();

            // Stimpak's own thread, so a retired client's stopping is seen even while the event loop
            // is busy — it can be the one waiting on it, when a "reconnect" command came in over chat.
            client.EventReceived += next =>
            {
                if (next is StageChanged { Stage: Stage.Disconnected } or SessionEnded)
                {
                    StoppedSc2Client(client);
                }
            };
            _ = SafeSc2ConsumeLoopAsync(client, _sc2ReceiveCts.Token);
        }

        _sc2Lease = lease;
        NoteSavedSignIn(profileId);
        Volatile.Write(ref _sc2LiveClient, client);
        RaiseActivityChanged();

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
                Channels = LastSc2Channels(profileId),
            });
        }
        catch (StimpakException ex)
        {
            // Stimpak only refuses a connect once the client's worker is gone — make a new one next time.
            LogError($"StarCraft II connect failed: {ex.Message}");
            _sc2ClientEnded = true;
            EndSc2Activity(client);
        }
    }

    /// <summary>
    /// Invigoration's own native client (the default, <see cref="NativeSc2ChatClient.Enabled"/>), or
    /// Stimpak's when INVIGORATION_STIMPAK=1 asks for the fallback. Either way credentials live under the bot's Battle.net profile.
    /// </summary>
    private ISc2ChatClient CreateSc2Client(string profileId)
    {
        if (NativeSc2ChatClient.Enabled && Config.Product == Protocol.BncsProduct.ScRemastered)
        {
            LogDebug("Using Invigoration's native StarCraft: Remastered connection.");
            return new NativeScrChatClient(profileId, Config.ScrGateway, Config.ScrCharacterName, LogDebug);
        }

        if (NativeSc2ChatClient.Enabled)
        {
            LogDebug("Using Invigoration's native StarCraft II connection.");
            return new NativeSc2ChatClient(profileId, LogDebug);
        }

        // ApplicationId is required but doesn't matter for us. CredentialPath overrides the
        // per-user cache location it would otherwise derive, since credential storage is already
        // fully owned by BattlenetCredentialProfileStore.
        return new StimpakSc2ChatClient(new StimpakClient(new StimpakClientOptions("cc.bnet.invigoration")
        {
            CredentialPath = BattlenetCredentialProfileStore.CredentialFilePath(profileId),
        }));
    }

    /// <summary>
    /// This bot's turn with its Battle.net profile's sign-in (see BattlenetSignInLease). Null when
    /// another bot or another copy of the app has it: that's logged, and any auto-reconnect stops.
    /// This bot's own previous client, retired mid-attempt and still finishing, is waited for instead.
    /// </summary>
    private async Task<BattlenetSignInLease?> AcquireSc2SignInAsync(string profileId, CancellationToken cancellationToken)
    {
        // Native connections sign in per game (see BattlenetSignInLease.NativeKey); Stimpak's share one.
        var leaseKey = NativeSc2ChatClient.Enabled ? BattlenetSignInLease.NativeKey(profileId, NativeProgram) : profileId;
        if (BattlenetSignInLease.TryAcquire(leaseKey, this, Config.DisplayName, out var lease, out var holder))
        {
            return lease;
        }

        if (holder is not null && ReferenceEquals(holder.Owner, this))
        {
            LogDebug("Waiting for the previous StarCraft II connection to finish closing...");
            try
            {
                // Counted as a connect in flight, so the bot looks busy and Disconnect can stop it.
                await TrackConnectAsync(
                    token => holder.Released.WaitAsync(Sc2WindDownTimeout + TimeSpan.FromSeconds(5), token),
                    cancellationToken).ConfigureAwait(false);
            }
            catch (TimeoutException)
            {
                // Released by then regardless; see ReleaseSc2ClientAsync.
            }

            if (BattlenetSignInLease.TryAcquire(leaseKey, this, Config.DisplayName, out lease, out holder))
            {
                return lease;
            }
        }

        var login = BattlenetCredentialProfileStore.Find(profileId)?.DisplayLabel ?? "this bot's Battle.net profile";
        var who = holder?.OwnerName ?? "another copy of Invigoration";
        var fix = holder is null ? "quit the other copy" : $"disconnect {holder.OwnerName}";
        LogError(
            $"The Battle.net login \"{login}\" is in use by {who}. One Battle.net login can only be connected {(NativeSc2ChatClient.Enabled ? "once per game" : "once")} at a time: " +
            $"{fix} first, or give this bot a Battle.net profile of its own (Edit Bot, then Battle.net Profile).");
        _logonRejection = $"its Battle.net login is in use by {who}.";
        return null;
    }

    /// <summary>
    /// Says whether this attempt has a saved sign-in to use, and how old it is. Every login replaces
    /// it, so its age is the time since the last one. If Battle.net then asks for a fresh sign-in, the
    /// warning names that age (see HandleSc2AuthenticationRequiredAsync). Only the file's timestamp is
    /// read, never its contents.
    /// </summary>
    private void NoteSavedSignIn(string profileId)
    {
        var path = NativeSc2ChatClient.Enabled
            ? BattlenetCredentialProfileStore.NativeCredentialFilePath(profileId, NativeProgram)
            : BattlenetCredentialProfileStore.CredentialFilePath(profileId);
        _sc2SavedSignInAge = File.Exists(path) && new FileInfo(path).Length > 0
            ? DateTime.UtcNow - File.GetLastWriteTimeUtc(path)
            : null;

        var login = BattlenetCredentialProfileStore.Find(profileId)?.DisplayLabel ?? "this bot";
        var message = _sc2SavedSignInAge is { } age
            ? $"Using the saved Battle.net sign-in for {login} (from {DescribeAge(age)} ago)."
            : $"No saved Battle.net sign-in for {login} yet. A Battle.net sign-in window will open.";
        if (IsReconnecting)
        {
            LogDebug(message);
        }
        else
        {
            LogInfo(message);
        }
    }

    /// <summary>The program code this bot's native client signs in as, which names its saved sign-in.</summary>
    private string NativeProgram => Config.Product == Protocol.BncsProduct.ScRemastered ? NativeScrChatClient.Program : NativeSc2ChatClient.Program;

    public static string DescribeAge(TimeSpan age) => age switch
    {
        { TotalDays: >= 2 } => $"{(int)age.TotalDays} days",
        { TotalHours: >= 2 } => $"{(int)age.TotalHours} hours",
        { TotalMinutes: >= 2 } => $"{(int)age.TotalMinutes} minutes",
        _ => $"{Math.Max(0, (int)age.TotalSeconds)} seconds",
    };

    /// <remarks>
    /// Reads until the client is disposed, which is after it's retired, so RetireSc2Client can see it
    /// stop. The token only reaches the sign-in: cancelling it closes a login window this client has open.
    /// </remarks>
    private async Task SafeSc2ConsumeLoopAsync(ISc2ChatClient client, CancellationToken signInCancellation)
    {
        try
        {
            await foreach (var next in client.ReadEventsAsync().ConfigureAwait(false))
            {
                await HandleSc2EventAsync(client, next, signInCancellation).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException)
        {
            // Nothing reads with a token now, but a handler's cancellation shouldn't end the process.
        }
        catch (Exception ex) when (ReferenceEquals(client, _sc2Client))
        {
            _sc2ClientEnded = true;
            LoseSc2Session($"StarCraft II connection lost: {ex.Message}", ex);
        }
        catch (Exception ex)
        {
            LogDebug($"A replaced StarCraft II client's events ended: {ex.Message}");
        }
        finally
        {
            // However the loop ended, this client's session is over — after the catch above has
            // already scheduled any reconnect, so the bot never looks idle in between.
            StoppedSc2Client(client);
            EndSc2Activity(client);
        }
    }

    private async Task HandleSc2EventAsync(ISc2ChatClient client, SC2Event next, CancellationToken cancellationToken = default)
    {
        client.People.Apply(next);

        // A client that's been replaced or disconnected can still have events queued up; acting
        // on them would close the current session's channel tabs or flip its status. Its stopping is
        // the one thing that matters: then it can be released (see RetireSc2Client).
        if (!ReferenceEquals(client, _sc2Client))
        {
            switch (next)
            {
                case StageChanged { Stage: Stage.Disconnected } or SessionEnded:
                    StoppedSc2Client(client);
                    break;

                // Battle.net turned down its saved sign-in after it was retired. Its Disconnect only
                // cancelled a sign-in already pending, so Stimpak would wait on this one forever; saying
                // no ends the attempt, and a Disconnected follows.
                case AuthenticationRequired auth:
                    try
                    {
                        client.CancelAuth(auth.AuthId);
                    }
                    catch (Exception ex) when (ex is StimpakException or ObjectDisposedException)
                    {
                        // Already over, or released.
                    }

                    break;
            }

            return;
        }

        switch (next)
        {
            case StageChanged { Stage: Stage.Connected }:
                // Chat is set up and the first channel joined (it arrives before this; any further
                // remembered channels arrive after, restored natively by Stimpak from
                // StimpakConnectOptions.Channels). That's "logged on" for SC2: an auto-reconnect
                // still running has nothing left to do.
                _sc2InChat = true;
                _autoReconnectCts?.Cancel();
                _connectedAt = DateTimeOffset.UtcNow;
                LogInfo("Connected to StarCraft II chat.");
                break;

            case StageChanged { Stage: Stage.Disconnected }:
                // Stimpak's one signal that an attempt or a session is over, however it ended —
                // including a live session dropped by the network or taken over by a sign-in
                // elsewhere, which Stimpak reports as SessionFailed then this (never SessionEnded;
                // that's only its worker dying). The client stays open for the next connect.
                LogDebug("StarCraft II stage: Disconnected");
                if (_sc2InChat || _sc2Channels.Count > 0)
                {
                    NoteSc2SessionLost();
                    LoseSc2Session("StarCraft II connection lost.", null);
                }

                EndSc2Activity(client);

                // A Disconnect that got in after this event was taken for the current client's is
                // waiting for a Disconnected that isn't coming: the worker is idle already.
                StoppedSc2Client(client);
                break;

            case StageChanged stage:
                LogDebug($"StarCraft II stage: {stage.Stage}");
                break;

            case AuthenticationRequired auth:
                // Stimpak has already deleted a saved sign-in Battle.net refused, before this arrives.
                // With none saved, NoteSavedSignIn already said a window would open.
                if (_sc2SavedSignInAge is { } age)
                {
                    LogWarning($"Battle.net didn't accept the saved sign-in (from {DescribeAge(age)} ago). Sign in again in the Battle.net window.");
                    _sc2SavedSignInAge = null;
                }

                await HandleSc2AuthenticationRequiredAsync(client, auth, cancellationToken).ConfigureAwait(false);
                break;

            case AccountConnected connected:
                // Ties this profile to the real signed-in BattleTag so it's identifiable if the
                // user has more than one Battle.net account — see BattlenetCredentialProfile
                // .DisplayLabel. Config.BattlenetCredentialProfileId is already guaranteed set by
                // now (ConnectSc2Async resolves it before Connect is called).
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

            case NativeChatEvent native:
                await HandleChatEvent(native.Event).ConfigureAwait(false);
                break;

            case NativeClubsEvent clubs:
                ClubsUpdated?.Invoke(clubs.Clubs);
                break;

            case NativeInvitationsEvent invitations:
                FriendInvitationsUpdated?.Invoke(invitations.Invitations);
                break;

            case NativeFriendsEvent friends:
                _sc2Friends.Clear();
                foreach (var friend in friends.Friends)
                {
                    _sc2Friends[friend.Account] = friend;
                }

                FriendsListUpdated?.Invoke(_sc2Friends.Values.ToList());
                break;

            case SessionFailed failed:
                // Why the attempt or session ended; Disconnected follows. A reconnect attempt
                // failing is expected along the way, so it's a detail rather than an error there.
                if (IsReconnecting && !_sc2InChat)
                {
                    LogDebug($"StarCraft II reconnect attempt failed: {failed.Message}");
                }
                else
                {
                    LogError($"StarCraft II session failed: {failed.Message}");
                }

                break;

            case SessionEnded:
                // Stimpak's worker is gone — in practice only if it crashed; a dropped or
                // taken-over session ends with a Disconnected stage instead (see above). This
                // client can't connect again, so the next connect makes a new one.
                _sc2ClientEnded = true;
                LoseSc2Session("StarCraft II session ended unexpectedly.", null);
                EndSc2Activity(client);
                StoppedSc2Client(client);
                break;

            // Not surfaced anywhere yet: roster snapshots (the UI binds Stimpak's own
            // PeopleRegistry.Channel(index) directly instead — see Sc2ChannelJoined),
            // group/party invitations.
            case RosterReceived or GroupInvitation or PartyInvitation or UnrecognisedEvent:
                break;
        }
    }

    /// <remarks>
    /// The token is the client's own event loop's, cancelled by Disconnect (and by removing the
    /// bot), which closes a login window still open — it used to stay up after the bot had moved
    /// on. A sign-in that's closed or can't be shown also stops auto-reconnect: every attempt would
    /// only ask again (a new window each time), and it's the user's call now.
    /// </remarks>
    private async Task HandleSc2AuthenticationRequiredAsync(ISc2ChatClient client, AuthenticationRequired auth, CancellationToken cancellationToken)
    {
        if (Sc2ChallengeHandler is null)
        {
            LogError("StarCraft II needs a sign-in, but no login window is available in this build.");
            _logonRejection = "StarCraft II needs a sign-in and none could be shown.";
            CancelSc2SignIn(client, auth);
            return;
        }

        try
        {
            var url = new Uri(auth.Url);
            var credential = await Sc2ChallengeHandler(url, cancellationToken).ConfigureAwait(false);
            client.SubmitAuth(auth.AuthId, System.Text.Encoding.UTF8.GetString(credential));
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            // Disconnect closed the window — the client is being torn down, nothing to answer.
        }
        catch (Exception ex)
        {
            LogError($"StarCraft II sign-in failed: {ex.Message}");
            _logonRejection = "the Battle.net sign-in wasn't completed.";
            CancelSc2SignIn(client, auth);
        }
    }

    /// <summary>
    /// Tells Stimpak nobody is going to finish this sign-in (the login window was closed, or there
    /// isn't one) — it otherwise waits for an answer forever, and the bot looks busy connecting
    /// with no way to try again. Stimpak ends the attempt with a Disconnected stage (see above).
    /// </summary>
    private void CancelSc2SignIn(ISc2ChatClient client, AuthenticationRequired auth)
    {
        // A login window left open across a Disconnect (or a removed bot) belongs to a client
        // that's already gone — disposed, maybe already replaced by a newer connect. There's no
        // sign-in left to cancel, and a disposed client throws rather than saying so.
        if (!ReferenceEquals(Volatile.Read(ref _sc2LiveClient), client))
        {
            return;
        }

        try
        {
            client.CancelAuth(auth.AuthId);
        }
        catch (Exception ex) when (ex is StimpakException or ObjectDisposedException)
        {
            // Already answered, cancelled, disconnected, or disposed in the meantime.
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
            // StarCraft II's limit is 255 characters; anything longer is cut there. (SC:R whispers
            // to a BattleTag go through Battle.net, not classic chat.)
            if (client is not NativeScrChatClient)
            {
                message = Invigoration.Sc2.Native.ChatCommands.CutToMessage(message);
            }

            try
            {
                client.SendWhisper(target, message);
            }
            catch (Exception ex) when (ex is StimpakException or ArgumentException)
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

        // StarCraft II chat has no slash commands beyond /w and /me (handled above), so anything
        // else would go out as a plain chat line. SC:R's server has its own; those pass through.
        // "//text" sends a line that starts with a slash.
        if (client is not NativeScrChatClient && body.StartsWith('/'))
        {
            if (!body.StartsWith("//", StringComparison.Ordinal))
            {
                LogError($"StarCraft II chat has no {body.Split(' ')[0]} command, so it wasn't sent. Start with // to send a line beginning with /.");
                return Task.CompletedTask;
            }

            body = body[1..];
        }

        // StarCraft II's limit is 255 characters, SC:R's classic chat 223; anything longer is cut there.
        body = client is NativeScrChatClient
            ? body[..Math.Min(body.Length, ChatLineSplitter.MaxLineLength)]
            : Invigoration.Sc2.Native.ChatCommands.CutToMessage(body);

        try
        {
            client.SendMessage(idx, body);
        }
        catch (Exception ex) when (ex is StimpakException or ArgumentException)
        {
            LogError($"Could not send: {ex.Message}");
        }

        return Task.CompletedTask;
    }

    /// <summary>
    /// Used to tear down the Stimpak client without telling anyone: no BncsDisconnected, no
    /// Sc2ChannelLeft for any open sub-tab. The underlying session genuinely did end (Dispose
    /// closes it at the protocol level, same as SessionEnded's real trigger), but the UI never
    /// found out — IsConnected stayed true, the status text stuck on "Connected", and every
    /// channel sub-tab stayed open and stale. Now mirrors what a server-driven SessionEnded
    /// already does: an explicit client.Disconnect() first (the native "end this session"
    /// call, not just freeing our local handle), then the same BncsDisconnected/Sc2ChannelLeft
    /// notifications DisconnectAsync's classic-BNCS path already fires via _bncs.Close().
    /// </summary>
    private Task DisconnectSc2Async()
    {
        _sc2InChat = false;
        _sc2QuickSessionLosses = 0;
        var hadClient = _sc2Client is not null;
        RetireSc2Client();
        CloseSc2Channels();

        if (hadClient)
        {
            BncsDisconnected?.Invoke(null);
        }

        return Task.CompletedTask;
    }

    /// <summary>
    /// Ends the current client for good: a login window it has open closes, its session is
    /// disconnected, and the client is released. That happens straight away if it was idle. If it was
    /// part-way through an attempt, it waits until Stimpak says it has stopped: Stimpak only takes a
    /// Disconnect once it's in chat, and until then it's still using the saved sign-in. The client's
    /// hold on the Battle.net sign-in goes with it, so a new connect on the same login waits for it
    /// rather than racing it. Tells nobody — the callers decide what the UI hears.
    /// </summary>
    private void RetireSc2Client()
    {
        _sc2ReceiveCts?.Cancel();
        _sc2ReceiveCts = null;
        var lease = Interlocked.Exchange(ref _sc2Lease, null);

        if (_sc2Client is not { } client)
        {
            lease?.Dispose();
            return;
        }

        _sc2Client = null;
        _sc2ClientProfileId = null;
        _sc2ClientEnded = false;

        // Registered before anything else, so its Disconnected can't go by unnoticed — including one
        // the event loop is handling for it right now as the current client's (see the Disconnected case).
        var stopped = _sc2WindingDown.GetOrAdd(client, _ => new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously)).Task;
        if (Interlocked.CompareExchange(ref _sc2LiveClient, null, client) == client)
        {
            RaiseActivityChanged();
        }
        else
        {
            // Idle: nothing to wait for.
            StoppedSc2Client(client);
        }

        try
        {
            client.Disconnect();
        }
        catch (Exception ex) when (ex is StimpakException or ObjectDisposedException)
        {
            // Its worker is already gone, so there's nothing to wait for.
            StoppedSc2Client(client);
        }

        _ = ReleaseSc2ClientAsync(client, stopped, lease);
    }

    /// <summary>A retired client said Disconnected, or its events ended: it has stopped.</summary>
    private void StoppedSc2Client(ISc2ChatClient client)
    {
        if (_sc2WindingDown.TryRemove(client, out var stopped))
        {
            stopped.TrySetResult();
        }
    }

    private async Task ReleaseSc2ClientAsync(ISc2ChatClient client, Task stopped, BattlenetSignInLease? lease)
    {
        try
        {
            await stopped.WaitAsync(Sc2WindDownTimeout).ConfigureAwait(false);
        }
        catch (TimeoutException)
        {
            LogDebug("A replaced StarCraft II connection didn't say it had stopped in time; releasing it anyway.");
        }
        finally
        {
            _sc2WindingDown.TryRemove(client, out _);
            client.Dispose();
            lease?.Dispose();
        }
    }

    /// <summary>
    /// Closes every channel sub-tab — without touching Config.Sc2LastChannels, which is exactly
    /// what the next connect restores.
    /// </summary>
    private void CloseSc2Channels()
    {
        var channelIndexes = _sc2Channels.Keys.ToList();
        _sc2Channels.Clear();
        _sc2ActiveChannelIndex = null;
        _sc2TriviaChannelIndex = null;
        foreach (var channelIndex in channelIndexes)
        {
            Sc2ChannelLeft?.Invoke(channelIndex);
        }
    }

    /// <summary>
    /// Counts a live session lost soon after getting into chat. A few of those in a row is another
    /// sign-in on the same Battle.net account taking the session each time: another bot, another copy
    /// of the app, or the game itself. Reconnecting would only take it back, and each would keep
    /// knocking the other off. Stopping sets a rejection, which MaybeScheduleAutoReconnect honours;
    /// a session that lasts resets the count, as does the user connecting or disconnecting.
    /// </summary>
    private void NoteSc2SessionLost()
    {
        if (!_sc2InChat)
        {
            return;
        }

        // A takeover partner comes back after its own reconnect delay plus a connect, so allow a
        // minute beyond this bot's delay.
        var window = TimeSpan.FromSeconds(Math.Max(1, Config.AutoReconnectDelaySeconds) + 60);
        _sc2QuickSessionLosses = DateTimeOffset.UtcNow - _connectedAt < window ? _sc2QuickSessionLosses + 1 : 0;
        if (_sc2QuickSessionLosses >= MaxSc2QuickSessionLosses)
        {
            _logonRejection =
                "StarCraft II chat keeps being taken over by another sign-in on this Battle.net account " +
                "(another bot, another copy of Invigoration, or the game itself).";
        }
    }

    /// <summary>A session that had made it into chat is gone: say so, close its tabs, and reconnect if that's on.</summary>
    private void LoseSc2Session(string message, Exception? ex)
    {
        _sc2InChat = false;
        CloseSc2Channels();
        LogError(message);
        BncsDisconnected?.Invoke(ex);
        MaybeScheduleAutoReconnect();
    }
}
