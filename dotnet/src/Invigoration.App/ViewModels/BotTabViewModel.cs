using System.Collections.ObjectModel;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.Models;
using Invigoration.Core;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;
using Invigoration.Core.Protocol;
using Stimpak;

namespace Invigoration.App.ViewModels;

/// <summary>One bot tab: wraps a BotEngine and projects its events onto observable collections for binding.</summary>
public partial class BotTabViewModel : ViewModelBase, IAsyncDisposable
{
    public BotEngine Engine { get; }

    public BotConfig Config => Engine.Config;

    public string Title => Config.DisplayName;

    /// <summary>Small game/client icon shown next to this tab's title — same icon key the Config window's own product picker uses (see BncsProduct.GetIconKey), just rendered smaller here.</summary>
    public Bitmap? TabIconImage => GameIconLoader.Get(BncsProduct.GetIconKey(Config.Product));

    /// <summary>The active bot's scheme-specific accent, for marking this tab as the open one and/or the chat input as focused.</summary>
    public IBrush HighlightBrush => new SolidColorBrush(
        Color.FromRgb(Engine.Palette.Highlight.R, Engine.Palette.Highlight.G, Engine.Palette.Highlight.B));

    /// <summary>The active bot's chat-log background, from its selected color scheme.</summary>
    public IBrush BackgroundBrush => new SolidColorBrush(
        Color.FromRgb(Engine.Palette.Background.R, Engine.Palette.Background.G, Engine.Palette.Background.B));

    /// <summary>The normal top-level tab header look — see GlobalWhispersTabViewModel's matching properties for why the Whispers pseudo-tab overrides both to stand out as a distinct, fixed utility tab.</summary>
    public double HeaderFontSize => 13;

    public IBrush HeaderForeground => Brushes.White;

    /// <summary>Whether this bot is the one actually visible right now — set by MainWindowViewModel.RecomputeActiveBot, which accounts for both being the selected top-level tab directly and being the selected member of a selected BotGroupTabViewModel. Setting this true clears HasUnread.</summary>
    [ObservableProperty]
    public partial bool IsActive { get; set; }

    partial void OnIsActiveChanged(bool value)
    {
        if (value)
        {
            HasUnread = false;
        }
    }

    /// <summary>A subtle "something happened while you weren't looking" flag for this bot's top-level tab — set on a new Talk/Emote/Broadcast line while !IsActive (see HandleChatEvent and OnChatMessage's multi-channel branch), cleared on becoming active.</summary>
    [ObservableProperty]
    public partial bool HasUnread { get; set; }

    public ObservableCollection<ChatLineViewModel> ChatLines { get; } = [];

    public ObservableCollection<ChannelUserViewModel> ChannelUsers { get; } = [];

    /// <summary>Whether this bot can be joined to several channels at once (SC2/SC:R/WC3:R) — gates the sub-tab UI. Classic BNCS/Chat-Telnet stay on the single flat ChatLines/ChannelUsers above.</summary>
    public bool SupportsMultiChannel => BncsProduct.IsStimpakBacked(Config.Product);

    /// <summary>One sub-tab per joined SC2/SC:R/WC3:R channel — see SupportsMultiChannel.</summary>
    public ObservableCollection<ChannelTabViewModel> Channels { get; } = [];

    [ObservableProperty]
    public partial ChannelTabViewModel? SelectedChannel { get; set; }

    /// <summary>The account's public-channel catalog (for the "join another channel" picker), sent once per SC2 session.</summary>
    public ObservableCollection<PublicChannel> AvailablePublicChannels { get; } = [];

    [ObservableProperty]
    public partial string JoinChannelName { get; set; } = "";

    public ObservableCollection<FriendEntryViewModel> Friends { get; } = [];

    /// <summary>One entry per peer this bot has whispered with, most-recently-active first — see UpsertWhisper. The only place a whisper's text is shown; it no longer also appears in the normal chat log (see HandleChatEvent's Whisper/WhisperSent cases).</summary>
    public ObservableCollection<WhisperThreadViewModel> WhisperThreads { get; } = [];

    /// <summary>Read-only snapshot of the shared roster for the Clan tab, filtered to formal members (IsClanMember) only — everyone else the bot has auto-tracked from chatting stays out of this tab, and only shows in the full Seen List window. Edits happen in the dedicated Clan Members window, opened via the "Manage Members..." button there.</summary>
    public ObservableCollection<ClanMemberViewModel> ClanRoster { get; } = [];

    /// <summary>Whether the current game's server pushes a friends list at all — false only for Diablo (1), which predates the feature entirely.</summary>
    public bool SupportsFriends => Invigoration.Core.Protocol.BncsProduct.SupportsFriendsList(Config.Product);

    /// <summary>
    /// Whether the Clan tab shows for this bot: the bot's own clan-management
    /// feature (roster/rank/alias/trivia-score — not Battle.net's native
    /// in-game clan protocol, which no product here speaks) has to be turned
    /// on in this bot's config, AND the shared roster needs at least one
    /// *formal* member (IsClanMember) — no point showing an always-empty tab
    /// before anyone's been explicitly added, and someone merely auto-tracked
    /// from chatting shouldn't count. The roster itself isn't per-product, so
    /// bots on different games (or eventually different platforms — SC2,
    /// SC:R) all see the same clan.
    /// </summary>
    public bool SupportsClan => Config.ClanFeatureEnabled && Invigoration.Core.Clan.ClanRosterStore.Members.Any(m => m.IsClanMember);

    [ObservableProperty]
    public partial string InputText { get; set; } = "";

    [ObservableProperty]
    public partial bool IsConnected { get; set; }

    [ObservableProperty]
    public partial string StatusText { get; set; } = "Disconnected";

    public bool DebugMode
    {
        get => Engine.DebugMode;
        set => Engine.DebugMode = value;
    }

    /// <summary>Swaps in an edited config (from the config window's Save) and refreshes anything derived from it, like the tab title.</summary>
    public void ApplyConfig(BotConfig newConfig)
    {
        Engine.Config = newConfig;
        OnPropertyChanged(nameof(Config));
        OnPropertyChanged(nameof(Title));
        OnPropertyChanged(nameof(TabIconImage));
        OnPropertyChanged(nameof(HighlightBrush));
        OnPropertyChanged(nameof(BackgroundBrush));
        OnPropertyChanged(nameof(SupportsFriends));
        OnPropertyChanged(nameof(SupportsClan));
    }

    public BotTabViewModel(BotEngine engine)
    {
        Engine = engine;
        ShowStartupBanner();
        Engine.Log += OnLog;
        Engine.SelfChatSent += OnSelfChatSent;
        Engine.ChatMessage += OnChatMessage;
        Engine.FriendsListUpdated += OnFriendsListUpdated;
        Engine.BncsConnected += () => Dispatcher.UIThread.Post(() =>
        {
            IsConnected = true;
            StatusText = "Connected";
        });
        Engine.BncsDisconnected += _ => Dispatcher.UIThread.Post(() =>
        {
            IsConnected = false;
            StatusText = "Disconnected";
            ChannelUsers.Clear();
            Friends.Clear();
        });
        Engine.Sc2ChannelJoined += OnSc2ChannelJoined;
        Engine.Sc2ChannelLeft += OnSc2ChannelLeft;
        Engine.Sc2ChannelJoinRejected += OnSc2ChannelJoinRejected;
        Engine.Sc2ChannelActionFailed += OnSc2ChannelActionFailed;
        Engine.Sc2PublicChannelsReceived += OnSc2PublicChannelsReceived;
        IconOverrideStore.OverridesChanged += OnIconOverrideChanged;
        Invigoration.Core.Clan.ClanRosterStore.RosterChanged += OnClanRosterChanged;
        RefreshClanRoster();
    }

    private void OnClanRosterChanged() => Dispatcher.UIThread.Post(() =>
    {
        OnPropertyChanged(nameof(SupportsClan));
        RefreshClanRoster();
    });

    private void RefreshClanRoster()
    {
        ClanRoster.Clear();
        foreach (var member in Invigoration.Core.Clan.ClanRosterStore.Members
                     .Where(m => m.IsClanMember)
                     .OrderBy(m => m.Rank).ThenBy(m => m.Name))
        {
            ClanRoster.Add(new ClanMemberViewModel(member));
        }
    }

    /// <summary>
    /// GameIconLoader's own bitmap cache already invalidates itself on an
    /// override change, but the already-bound ChannelUserViewModel/
    /// FriendEntryViewModel rows on screen won't re-pull it unless something
    /// tells them to — otherwise a swapped icon only shows up after a
    /// reconnect rebuilds the list from scratch. This makes it immediate.
    /// Also re-raises TabIconImage itself — this bot's own tab-strip icon uses
    /// the exact same key (e.g. applying the Battle.net 2.0 icon set
    /// overrides "sc"/"war3", the same keys BncsProduct.GetIconKey resolves
    /// SC:Remastered/WC3:Reforged to) and was missing this notification
    /// entirely, so a SC:R/WC3:R bot's tab kept showing the classic icon
    /// until the app restarted even though the override had actually applied.
    /// </summary>
    private void OnIconOverrideChanged(string key) => Dispatcher.UIThread.Post(() =>
    {
        foreach (var user in ChannelUsers)
        {
            user.RefreshIcons();
        }

        foreach (var friend in Friends)
        {
            friend.RefreshIcon();
        }

        if (key.Equals(BncsProduct.GetIconKey(Config.Product), StringComparison.OrdinalIgnoreCase))
        {
            OnPropertyChanged(nameof(TabIconImage));
        }
    });

    /// <summary>
    /// Restores the VB6 original's Form_Load ASCII-art banner - a bunny made
    /// of parentheses plus "Invigoration Beta bunny" in red/green - shown
    /// once when a bot tab opens, as a nod to this project's long-running
    /// beta status. Ported from frmMain.frm's AddChat calls; the colored
    /// "Beta"/"bunny" words reuse the same inline color-code marker
    /// (U+00A0 + letter) ChatColorFormatter already parses everywhere else.
    /// </summary>
    private void ShowStartupBanner()
    {
        var p = Engine.Palette;
        const string separator = "---------------------------------------------------";
        const char marker = ' ';
        var bunnyLine = $"Invigoration {marker}rBeta {marker}gbunny";

        ChatLines.Add(new ChatLineViewModel(separator, p.Highlight));
        ChatLines.Add(new ChatLineViewModel("()()", p.Info));
        ChatLines.Add(new ChatLineViewModel("(--)", p.Info));
        ChatLines.Add(new ChatLineViewModel("(')(')", p.Info));
        ChatLines.Add(new ChatLineViewModel(ChatColorFormatter.Parse(bunnyLine, p.Channel, p)));
        ChatLines.Add(new ChatLineViewModel("C#/.NET port -- still in beta", p.Debug));
        ChatLines.Add(new ChatLineViewModel(separator, p.Info));
    }

    [RelayCommand]
    private async Task ConnectAsync()
    {
        try
        {
            StatusText = "Connecting...";
            await Engine.ConnectAsync();
        }
        catch (Exception ex)
        {
            ChatLines.Add(new ChatLineViewModel($"Connect failed: {ex.Message}", Engine.Palette.Error));
            StatusText = "Disconnected";
        }
    }

    [RelayCommand]
    private Task DisconnectAsync() => Engine.DisconnectAsync();

    /// <summary>Diablo II's server doesn't push status updates automatically, so a manual refresh is the only way to see current friend status there.</summary>
    [RelayCommand]
    private Task RefreshFriendsAsync() => Engine.RequestFriendsListAsync();

    [RelayCommand]
    private async Task SendAsync()
    {
        var text = InputText;
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        InputText = "";

        // Only "/" runs a local command now — the configured Trigger character no longer does
        // (see BotEngine.Commands.cs), so anything typed locally that starts with it is just
        // sent as ordinary chat text below, same as any other message. "//" still escapes a
        // leading slash: sends the rest verbatim as a real chat message instead of intercepting
        // it as a local command — lets you test another bot's slash command (e.g. "//join foo")
        // from this bot's own tab as if you were just another channel member.
        if (text.Length > 0 && text[0] == '/')
        {
            if (text.Length > 1 && text[1] == '/')
            {
                await Engine.SendChatCommandAsync(text[1..]);
            }
            else
            {
                await Engine.RunLocalCommandAsync(text);
            }
        }
        else
        {
            // Engine.SendChatCommandAsync echoes non-command text into the
            // log itself (Battle.net doesn't echo a client's own messages).
            await Engine.SendChatCommandAsync(text);
        }
    }

    private void OnLog(IReadOnlyList<ChatLogSegment> segments) =>
        Dispatcher.UIThread.Post(() => ChatLines.Add(new ChatLineViewModel(segments)));

    /// <summary>A message this bot itself just sent — routed to wherever it actually went (the active sub-tab for a multi-channel bot, the flat log otherwise), with the same speaker-icon resolution a real Talk event gets.</summary>
    private void OnSelfChatSent(IReadOnlyList<ChatLogSegment> segments) => Dispatcher.UIThread.Post(() =>
    {
        if (SupportsMultiChannel)
        {
            SelectedChannel?.ChatLines.Add(new ChatLineViewModel(segments, ResolveSc2UserIcon()));
        }
        else
        {
            ChatLines.Add(new ChatLineViewModel(segments, ResolveUserIcon(Config.Username)));
        }
    });

    private static bool IsUnreadWorthy(ChatEventType type) => type is ChatEventType.Talk or ChatEventType.Emote or ChatEventType.Broadcast;

    private void OnChatMessage(ChatEvent e) => Dispatcher.UIThread.Post(() =>
    {
        if (SupportsMultiChannel && e.ChannelIndex is { } channelIndex)
        {
            var channel = Channels.FirstOrDefault(c => c.ChannelIndex == channelIndex);
            channel?.HandleChatEvent(e, Engine.Palette, ResolveSc2UserIcon());
            if (IsUnreadWorthy(e.Type))
            {
                if (channel is not null && channel != SelectedChannel)
                {
                    channel.HasUnread = true;
                }

                if (!IsActive)
                {
                    HasUnread = true;
                }
            }

            return;
        }

        HandleChatEvent(e);
    });

    private void OnSc2ChannelJoined(byte channelIndex, ChatChannel channel, ObservableCollection<Person> users) =>
        Dispatcher.UIThread.Post(() =>
        {
            // Defensive: a duplicate Joined for an index this UI already has a tab for would
            // otherwise add a second ChannelTabViewModel with the same ChannelIndex, and every
            // later lookup-by-index (leave, active-channel tracking) only ever finds the first
            // match — the second becomes an orphaned, stuck tab with no way to close it.
            if (Channels.Any(c => c.ChannelIndex == channelIndex))
            {
                return;
            }

            var tab = new ChannelTabViewModel(channelIndex, channel, users);
            Channels.Add(tab);
            SelectedChannel ??= tab;
        });

    private void OnSc2ChannelLeft(byte channelIndex) => Dispatcher.UIThread.Post(() =>
    {
        var tab = Channels.FirstOrDefault(c => c.ChannelIndex == channelIndex);
        if (tab is null)
        {
            return;
        }

        Channels.Remove(tab);
        if (SelectedChannel == tab)
        {
            SelectedChannel = Channels.Count > 0 ? Channels[0] : null;
        }
    });

    private void OnSc2ChannelJoinRejected(string reason) => Dispatcher.UIThread.Post(() =>
        (SelectedChannel?.ChatLines ?? ChatLines).Add(new ChatLineViewModel(reason, Engine.Palette.Error)));

    private void OnSc2ChannelActionFailed(string reason) => Dispatcher.UIThread.Post(() =>
        (SelectedChannel?.ChatLines ?? ChatLines).Add(new ChatLineViewModel(reason, Engine.Palette.Error)));

    private void OnSc2PublicChannelsReceived(IReadOnlyList<ChatChannel> channels) => Dispatcher.UIThread.Post(() =>
    {
        AvailablePublicChannels.Clear();
        foreach (var channel in channels.OfType<PublicChannel>())
        {
            AvailablePublicChannels.Add(channel);
        }
    });

    partial void OnSelectedChannelChanged(ChannelTabViewModel? value)
    {
        if (value is not null)
        {
            Engine.SetActiveSc2Channel(value.ChannelIndex);
            value.HasUnread = false;
        }
    }

    [RelayCommand]
    private void LeaveChannel(ChannelTabViewModel tab) => Engine.LeaveSc2Channel(tab.ChannelIndex);

    [RelayCommand]
    private void JoinPublicChannel(PublicChannel channel) => Engine.TryJoinSc2PublicChannel(channel.Id);

    [RelayCommand]
    private void JoinPrivateChannel()
    {
        if (string.IsNullOrWhiteSpace(JoinChannelName))
        {
            return;
        }

        if (Engine.TryJoinSc2PrivateChannel(JoinChannelName))
        {
            JoinChannelName = "";
        }
    }

    /// <summary>Battle.net's own system account for account-notification whispers (e.g. "you have unread mail") — not a real user, pure noise for a bot. Ignored only for incoming whispers; a deliberately-sent outgoing one (unlikely, but conceivable) isn't suppressed.</summary>
    private const string IgnoredWhisperSender = "# Email Service #";

    /// <summary>Finds or creates an empty thread for a peer with no message appended — for the right-click "Whisper" action (BotTabView.axaml.cs/MainWindowViewModel.FocusWhisperThread), which just needs somewhere to open a compose box, not a logged message.</summary>
    public WhisperThreadViewModel GetOrCreateWhisperThread(string peer)
    {
        var thread = WhisperThreads.FirstOrDefault(t => t.Peer == peer);
        if (thread is null)
        {
            thread = new WhisperThreadViewModel(this, peer);
            WhisperThreads.Insert(0, thread);
        }

        return thread;
    }

    /// <summary>Finds or creates the thread for a peer, appends the message, and bumps it to the top of WhisperThreads (most-recently-active first) — the single entry point both incoming Whisper and outgoing WhisperSent events go through, see HandleChatEvent.</summary>
    private WhisperThreadViewModel? UpsertWhisper(string peer, string text, bool incoming, ChatPalette palette)
    {
        if (incoming && string.Equals(peer, IgnoredWhisperSender, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var thread = WhisperThreads.FirstOrDefault(t => t.Peer == peer);
        if (thread is null)
        {
            thread = new WhisperThreadViewModel(this, peer);
            WhisperThreads.Insert(0, thread);
        }
        else
        {
            var currentIndex = WhisperThreads.IndexOf(thread);
            if (currentIndex != 0)
            {
                WhisperThreads.Move(currentIndex, 0);
            }
        }

        var lineText = incoming ? $"{peer}: {text}" : $"You: {text}";
        thread.Messages.Add(new ChatLineViewModel(lineText, incoming ? palette.Whisper : palette.SelfUserName));
        thread.LastActivityUtc = DateTime.UtcNow;
        if (incoming)
        {
            thread.HasUnread = true;
        }

        return thread;
    }

    /// <summary>Sends a whisper thread's DraftText to its peer through this bot's own engine — works for both classic BNCS (server-parsed "/w") and Stimpak-backed products (BotEngine.Sc2.cs intercepts the same "/w " convention and routes it to Stimpak's dedicated whisper API instead).</summary>
    [RelayCommand]
    private async Task SendWhisperAsync(WhisperThreadViewModel thread)
    {
        var text = thread.DraftText.Trim();
        if (text.Length == 0)
        {
            return;
        }

        thread.DraftText = "";
        await Engine.SendChatCommandAsync($"/w {thread.Peer} {text}");
    }

    /// <summary>
    /// Reconciles the Friends collection with the engine's current list by
    /// account name rather than clearing and rebuilding, so an in-place
    /// status update (the common case — SID_FRIENDSUPDATE) doesn't disturb
    /// list selection/scroll position. Also applies SID_FRIENDSPOSITION
    /// reordering, since the engine's list is already in server order.
    /// </summary>
    private void OnFriendsListUpdated(IReadOnlyList<FriendEntry> entries) => Dispatcher.UIThread.Post(() =>
    {
        for (var i = Friends.Count - 1; i >= 0; i--)
        {
            if (entries.All(e => e.Account != Friends[i].Account))
            {
                Friends.RemoveAt(i);
            }
        }

        for (var i = 0; i < entries.Count; i++)
        {
            var entry = entries[i];
            var friend = Friends.FirstOrDefault(f => f.Account == entry.Account);
            if (friend is null)
            {
                friend = new FriendEntryViewModel(entry.Account);
                Friends.Insert(Math.Min(i, Friends.Count), friend);
            }
            else
            {
                var currentIndex = Friends.IndexOf(friend);
                if (currentIndex != i)
                {
                    Friends.Move(currentIndex, i);
                }
            }

            friend.Status = entry.Status;
            friend.Location = entry.Location;
            friend.ProductCode = entry.ProductCode;
            friend.LocationName = entry.LocationName;
        }

        // Online-first, otherwise stable (OrderByDescending doesn't reorder two friends that
        // are both online, or both offline, relative to each other — so within each group this
        // keeps the server's own position order from the loop above). Applied as a sequence of
        // in-place Moves, not a rebuild, for the same selection/scroll-preserving reason the
        // reconciliation above is structured this way.
        var sorted = Friends.OrderByDescending(f => f.IsOnline).ToList();
        for (var i = 0; i < sorted.Count; i++)
        {
            var currentIndex = Friends.IndexOf(sorted[i]);
            if (currentIndex != i)
            {
                Friends.Move(currentIndex, i);
            }
        }
    });

    private void HandleChatEvent(ChatEvent e)
    {
        var palette = Engine.Palette;
        if (IsUnreadWorthy(e.Type) && !IsActive)
        {
            HasUnread = true;
        }

        switch (e.Type)
        {
            case ChatEventType.Channel:
                ChannelUsers.Clear();
                ChatLines.Add(new ChatLineViewModel($"*** Joined channel: {e.Text}", palette.Channel));
                break;

            case ChatEventType.ShowUser:
            case ChatEventType.Join:
                UpsertUser(e);
                if (e.Type == ChatEventType.Join)
                {
                    ChatLines.Add(new ChatLineViewModel($"*** {e.Username} has joined the channel.", palette.Gray));
                }

                break;

            case ChatEventType.Leave:
                var leaving = ChannelUsers.FirstOrDefault(u => u.Username == e.Username);
                if (leaving is not null)
                {
                    ChannelUsers.Remove(leaving);
                }

                ChatLines.Add(new ChatLineViewModel($"*** {e.Username} has left the channel.", palette.Gray));
                break;

            case ChatEventType.UserFlags:
                UpsertUser(e);
                break;

            case ChatEventType.Talk:
                ChatLines.Add(new ChatLineViewModel(BuildUserLine(e.Username, e.Text, e.Flags, palette), ResolveUserIcon(e.Username)));
                break;

            case ChatEventType.Emote:
                ChatLines.Add(new ChatLineViewModel($"<{e.Username} {e.Text}>", palette.GetEmoteColor(e.Flags), ResolveUserIcon(e.Username)));
                break;

            case ChatEventType.Whisper:
                UpsertWhisper(e.Username, e.Text, incoming: true, palette);
                break;

            case ChatEventType.WhisperSent:
                UpsertWhisper(e.Username, e.Text, incoming: false, palette);
                break;

            case ChatEventType.Info:
                ChatLines.Add(new ChatLineViewModel(e.Text, palette.Info));
                break;

            case ChatEventType.Error:
                ChatLines.Add(new ChatLineViewModel(e.Text, palette.Error));
                break;

            case ChatEventType.Broadcast:
                ChatLines.Add(new ChatLineViewModel($"[Broadcast]: {e.Text}", palette.Debug));
                break;
        }
    }

    /// <summary>
    /// Adds a newly-seen user, or updates an already-tracked one's flags/ping/statstring —
    /// either way, (re)positions them via InsertUserSorted so a promotion/demotion (this same
    /// method handles ChatEventType.UserFlags too) actually moves them in the list, matching
    /// classic Battle.net's own "moderators float to the top" behavior instead of leaving
    /// everyone frozen in original join order regardless of rank changes.
    /// </summary>
    private void UpsertUser(ChatEvent e)
    {
        var user = ChannelUsers.FirstOrDefault(u => u.Username == e.Username);
        if (user is not null)
        {
            ChannelUsers.Remove(user);
        }
        else
        {
            user = new ChannelUserViewModel(e.Username) { UseClassicIconStyle = Config.ClassicUserIconStyle };
        }

        user.Flags = e.Flags;
        user.Ping = e.Ping;
        if (e.Text.Length > 0)
        {
            user.StatString = e.Text;
        }

        InsertUserSorted(user);
    }

    /// <summary>Privileged users (see ChatIcon.IsPrivileged — Blizzard/Admin/Operator/Speaker) sort to the top, in their own arrival order; everyone else keeps arriving at the bottom, in theirs — the classic Battle.net "moderators, then users, each by join time" ordering.</summary>
    private void InsertUserSorted(ChannelUserViewModel user)
    {
        if (!ChatIcon.IsPrivileged(user.Flags))
        {
            ChannelUsers.Add(user);
            return;
        }

        var insertIndex = 0;
        while (insertIndex < ChannelUsers.Count && ChatIcon.IsPrivileged(ChannelUsers[insertIndex].Flags))
        {
            insertIndex++;
        }

        ChannelUsers.Insert(insertIndex, user);
    }

    /// <summary>
    /// Flips Config.ClassicUserIconStyle and pushes the new value into every already-tracked row
    /// so the Users list updates immediately — called from the right-click "Classic Icon Style"
    /// menu item (BotTabView.axaml.cs), not the Config window, per explicit request. Persisted
    /// the same way every other BotConfig field is: whenever SaveAll next runs (window close, or
    /// after the Config window itself is saved), not immediately here.
    /// </summary>
    public void ToggleClassicUserIconStyle()
    {
        Config.ClassicUserIconStyle = !Config.ClassicUserIconStyle;
        foreach (var user in ChannelUsers)
        {
            user.UseClassicIconStyle = Config.ClassicUserIconStyle;
        }
    }

    private static IReadOnlyList<ChatLogSegment> BuildUserLine(string username, string text, uint flags, ChatPalette palette)
    {
        var segments = new List<ChatLogSegment> { new(palette.GetUserNameColor(flags), $"{username}: ") };
        segments.AddRange(ChatColorFormatter.Parse(text, palette.GetChatColor(flags), palette));
        return segments;
    }

    /// <summary>Classic BNCS speaker icon, from whatever statstring the userlist already tracked for them — see BotConfig.ShowUserIconsInChat. Null once the toggle is off, or for a name with no tracked statstring yet (e.g. a whisper-only stranger who's never actually spoken in-channel).</summary>
    private Bitmap? ResolveUserIcon(string username)
    {
        if (!Config.ShowUserIconsInChat)
        {
            return null;
        }

        var statString = ChannelUsers.FirstOrDefault(u => u.Username == username)?.StatString;
        if (string.IsNullOrEmpty(statString))
        {
            return null;
        }

        var key = ChatIcon.GetProductIconKey(statString);
        return string.IsNullOrEmpty(key) ? null : GameIconLoader.Get(key);
    }

    /// <summary>Stimpak-backed (SC2/SC:R/WC3:R) speaker icon — every speaker gets this bot's own product icon, since Stimpak's roster carries no per-user product field to tell them apart by (see BotConfig.ShowUserIconsInChat's remarks).</summary>
    private Bitmap? ResolveSc2UserIcon() =>
        Config.ShowUserIconsInChat ? GameIconLoader.Get(BncsProduct.GetIconKey(Config.Product)) : null;

    public ValueTask DisposeAsync()
    {
        Engine.Log -= OnLog;
        Engine.SelfChatSent -= OnSelfChatSent;
        Engine.ChatMessage -= OnChatMessage;
        Engine.FriendsListUpdated -= OnFriendsListUpdated;
        Engine.Sc2ChannelJoined -= OnSc2ChannelJoined;
        Engine.Sc2ChannelLeft -= OnSc2ChannelLeft;
        Engine.Sc2ChannelJoinRejected -= OnSc2ChannelJoinRejected;
        Engine.Sc2ChannelActionFailed -= OnSc2ChannelActionFailed;
        Engine.Sc2PublicChannelsReceived -= OnSc2PublicChannelsReceived;
        IconOverrideStore.OverridesChanged -= OnIconOverrideChanged;
        Invigoration.Core.Clan.ClanRosterStore.RosterChanged -= OnClanRosterChanged;
        return Engine.DisposeAsync();
    }
}
