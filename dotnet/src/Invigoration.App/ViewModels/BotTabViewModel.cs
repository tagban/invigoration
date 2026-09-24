using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.Threading;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.Models;
using Invigoration.Core;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;
using Invigoration.Core.Sc2;
using Invigoration.Core.Discord;
using Invigoration.Core.Protocol;
using Stimpak;

namespace Invigoration.App.ViewModels;

/// <summary>One bot tab: wraps a BotEngine and projects its events onto observable collections for binding.</summary>
public partial class BotTabViewModel : ViewModelBase, IAsyncDisposable, IThemedSurface, IWhisperHost
{
    public BotEngine Engine { get; }

    /// <summary>
    /// This bot's theme (BotConfig.ThemeId, resolved through ThemeLibrary) as the brushes and layout
    /// coordinates BotTabView draws with. Rebuilt whenever it could have changed: the bot's config
    /// is replaced, or a custom theme is saved or deleted in the theme manager.
    /// </summary>
    [ObservableProperty]
    public partial ChatThemeViewModel Theme { get; set; } = new(ThemeLibrary.BuiltIns[0]);

    partial void OnThemeChanged(ChatThemeViewModel value)
    {
        value.ChannelName = CurrentChannelName;
        foreach (var user in ChannelUsers)
        {
            user.UseD2Layout = value.UsesCharacterDock;
        }
    }

    private void RefreshTheme() => Theme = new ChatThemeViewModel(ThemeLibrary.ResolveFor(Config));

    private void OnThemesChanged() => Dispatcher.UIThread.Post(RefreshTheme);

    private void OnD2EquipmentChanged() => Dispatcher.UIThread.Post(() =>
    {
        foreach (var user in ChannelUsers)
        {
            user.RefreshD2Data();
        }
    });

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

    /// <summary>A subtle "something happened while you weren't looking" flag for this bot's top-level tab — set on a new Talk/Emote/Broadcast line while !IsActive (see HandleChatEvent and ProcessChatEvent's multi-channel branch), cleared on becoming active.</summary>
    [ObservableProperty]
    public partial bool HasUnread { get; set; }

    public ObservableCollection<ChatLineViewModel> ChatLines { get; } = [];

    /// <summary>Fired after ChatLineTrimmer trims old lines off the front of ChatLines — see its own remarks. The view rebuilds its Inlines from the (now bounded) collection in response.</summary>
    public event Action? ChatLinesTrimmed;

    public ObservableCollection<ChannelUserViewModel> ChannelUsers { get; } = [];

    /// <summary>
    /// Username -> row index for ChannelUsers, kept in sync at every add/remove/clear site below.
    /// A plain ChannelUsers.FirstOrDefault(u => u.Username == ...) scan is O(n) per lookup, which
    /// made a burst of many users joining/leaving a channel at once (e.g. a mass-join flood) cost
    /// O(n²) overall — confirmed live as the cause of the UI falling badly behind during one.
    /// </summary>
    private readonly Dictionary<string, ChannelUserViewModel> _channelUsersByName = new();

    /// <summary>The classic-BNCS channel this bot is currently in, from the server's own ChatEventType.Channel notification — "" before joining one (or after a disconnect). Only meaningful for a non-SupportsMultiChannel bot; see UsersTabHeader.</summary>
    [ObservableProperty]
    public partial string CurrentChannelName { get; set; } = "";

    /// <summary>"Users" Tab header text — just "Users" until a channel name is known, then "<channel> (<count>)" so the tab itself answers "how many people are in here" without opening it. Recomputed (via the OnPropertyChanged calls scattered below) whenever anything it reads changes: CurrentChannelName, ChannelUsers' count, SelectedChannel, or the selected channel's own Users count.</summary>
    public string UsersTabHeader
    {
        get
        {
            if (SupportsMultiChannel)
            {
                return SelectedChannel is { } channel ? $"{channel.Title} ({channel.Users.Count})" : "Users";
            }

            return string.IsNullOrEmpty(CurrentChannelName) ? "Users" : $"{CurrentChannelName} ({ChannelUsers.Count})";
        }
    }

    /// <summary>Whether this bot can be joined to several channels at once (SC2/SC:R/WC3:R) — gates the sub-tab UI. Classic BNCS/Chat-Telnet stay on the single flat ChatLines/ChannelUsers above.</summary>
    public bool SupportsMultiChannel => BncsProduct.IsStimpakBacked(Config.Product);

    /// <summary>
    /// StarCraft: Remastered on the native connection: one channel at a time, like classic
    /// Battle.net, so its channel shows without tabs and joining another switches to it. Stimpak's
    /// SC:R is really SC2 and keeps the tabs.
    /// </summary>
    /// <summary>The user's answer to the one-time StarCraft II portraits offer.</summary>
    public enum Sc2PortraitsAnswer
    {
        Download,
        NotNow,
        Never,
    }

    /// <summary>Asks the user whether to download the StarCraft II portraits; set by the main window.</summary>
    public static Func<Task<Sc2PortraitsAnswer>>? AskToDownloadSc2Portraits { get; set; }

    private static bool _sc2PortraitsOffered;

    private int _portraitRedrawQueued;

    /// <summary>Portraits arrive a few at a time; redraw once per burst, a quarter second after the first.</summary>
    private void OnPortraitsChanged()
    {
        if (!SupportsMultiChannel || Interlocked.Exchange(ref _portraitRedrawQueued, 1) == 1)
        {
            return;
        }

        DispatcherTimer.RunOnce(() =>
        {
            Interlocked.Exchange(ref _portraitRedrawQueued, 0);
            IconVersion++;
        }, TimeSpan.FromMilliseconds(250));
    }

    /// <summary>
    /// The first time a StarCraft II bot connects (natively) without the portraits, offers to
    /// download them, once per run. They're Blizzard's art, so they aren't shipped with the app.
    /// </summary>
    private async Task OfferSc2PortraitsAsync()
    {
        if (Config.Product != BncsProduct.Sc2 || !Invigoration.Core.Sc2.NativeSc2ChatClient.Enabled
            || _sc2PortraitsOffered || Sc2PortraitStore.Declined || Sc2PortraitStore.IsDownloaded
            || AskToDownloadSc2Portraits is not { } ask)
        {
            return;
        }

        _sc2PortraitsOffered = true;
        switch (await ask())
        {
            case Sc2PortraitsAnswer.Never:
                Sc2PortraitStore.Decline();
                return;
            case Sc2PortraitsAnswer.NotNow:
                return;
        }

        void Say(string text) => (SelectedChannel?.ChatLines ?? ChatLines).Add(new ChatLineViewModel(text, Engine.Palette.Info));
        Say("Downloading StarCraft II portraits (about 21 MB)...");
        try
        {
            await Sc2PortraitStore.DownloadAsync(null, CancellationToken.None);
            Say("StarCraft II portraits downloaded.");
        }
        catch (Exception ex) when (ex is HttpRequestException or IOException or TaskCanceledException or UnauthorizedAccessException)
        {
            Say($"Couldn't download the StarCraft II portraits: {ex.Message} You'll be asked again next time.");
            _sc2PortraitsOffered = false;
        }
    }

    /// <summary>The SC2/SC:R user list's picture size: compact by default; larger in SC2's Full view (BotConfig.FullUserListPortraits).</summary>
    public double PortraitSize => ShowsMemberDetail ? 36 : 20;

    /// <summary>SC2's Full user list: a light detail line (BattleTag, status) under each name.</summary>
    public bool ShowsMemberDetail => IsSc2 && Config.FullUserListPortraits;

    public void SetUserListFull(bool full)
    {
        Config.FullUserListPortraits = full;
        OnPropertyChanged(nameof(PortraitSize));
        OnPropertyChanged(nameof(ShowsMemberDetail));
    }

    /// <summary>The user list's name size: a notch smaller for SC2, whose clan tags make names long. SC:R has none.</summary>
    public double UserListNameFontSize => IsSc2 ? 12 : 14;

    /// <summary>Whether this bot shows SC2 portraits (as opposed to SC:R's classic game icons, which follow icon sets).</summary>
    public bool IsSc2 => Config.Product == BncsProduct.Sc2;

    /// <summary>The "Pictures in Chat" choice: off, compact (a small icon before the line) or full (SC2: a portrait beside name and text).</summary>
    public void SetChatPictures(bool show, bool full)
    {
        Config.ShowUserIconsInChat = show;
        Config.FullChatPortraits = full;
    }

    /// <summary>Bumped when icons change, so the SC2/SC:R user rows (Stimpak's Person, which can't be told) redraw theirs.</summary>
    [ObservableProperty]
    public partial int IconVersion { get; set; }

    public bool IsSingleChannel => Config.Product == BncsProduct.ScRemastered && Invigoration.Core.Sc2.NativeSc2ChatClient.Enabled;

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
    [NotifyPropertyChangedFor(nameof(CanConnect), nameof(CanDisconnect), nameof(StatusDotBrush))]
    public partial bool IsConnected { get; set; }

    /// <summary>The engine's own "nothing connected, connecting, or waiting to reconnect" (BotEngine.IsIdle), refreshed on the UI thread whenever it says to look again.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanConnect), nameof(CanDisconnect))]
    public partial bool IsIdle { get; set; } = true;

    /// <summary>An auto-reconnect sitting out its delay with nothing in flight (BotEngine.IsWaitingToReconnect) — the bot's disconnected, and Connect skips the wait.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanConnect), nameof(ConnectButtonText), nameof(ConnectButtonTip))]
    public partial bool IsWaitingToReconnect { get; set; }

    /// <summary>"Reconnect Now" while an auto-reconnect is only waiting, since that's what a click does — skips the wait.</summary>
    public string ConnectButtonText => IsWaitingToReconnect ? "Reconnect Now" : "Connect";

    public string? ConnectButtonTip => IsWaitingToReconnect ? "Auto-reconnect is waiting between attempts — connect right away instead" : null;

    private static readonly IBrush ConnectedDot = new SolidColorBrush(Color.FromRgb(0x3F, 0xB9, 0x50));
    private static readonly IBrush BusyDot = new SolidColorBrush(Color.FromRgb(0xD2, 0x99, 0x22));

    /// <summary>Green once connected; amber while connecting or reconnecting (the only other time the dot shows).</summary>
    public IBrush StatusDotBrush => IsConnected ? ConnectedDot : BusyDot;

    /// <summary>
    /// Whether Connect makes sense right now — what the Connect button's visibility and the Bot
    /// menu's Connect item follow: the bot is disconnected with nothing under way, or only
    /// counting down to a reconnect. Also requires !IsConnected so a path the engine's own check
    /// ever missed still can't show Connect on a bot that's plainly online.
    /// </summary>
    public bool CanConnect => !IsConnected && (IsIdle || IsWaitingToReconnect);

    /// <summary>Anything to stop — a live connection, an attempt still under way, or a reconnect counting down.</summary>
    public bool CanDisconnect => IsConnected || !IsIdle;

    [ObservableProperty]
    public partial string StatusText { get; set; } = "Disconnected";

    /// <summary>Not saved — like /debug, it's for chasing a problem in this session. The Bot menu's checkmark follows it through PropertyChanged whichever way it's switched.</summary>
    public bool DebugMode
    {
        get => Engine.DebugMode;
        set
        {
            Engine.DebugMode = value;
            OnPropertyChanged();
        }
    }

    /// <summary>Everything a lost connection takes with it: the connected flag and whoever was in the channel or on the friends list.</summary>
    private void MarkDisconnected()
    {
        IsConnected = false;
        ChannelUsers.Clear();
        _channelUsersByName.Clear();
        CurrentChannelName = "";
        Friends.Clear();
    }

    /// <summary>Re-reads the engine's connection state and words the status to match: an attempt or countdown in progress says so rather than looking like a plain "Disconnected" with no Connect button beside it.</summary>
    private void RefreshConnectionState()
    {
        IsIdle = Engine.IsIdle;
        IsWaitingToReconnect = Engine.IsWaitingToReconnect;

        // IsConnected is only ever cleared by a BncsDisconnected, and a few ends raise none — an
        // SC2 session lost to the network, say, which Stimpak only reports as a stage change. An
        // engine with nothing up can't be online, so its own state wins over the latched flag.
        if (IsConnected && (IsIdle || IsWaitingToReconnect))
        {
            MarkDisconnected();
        }

        StatusText = IsConnected ? "Connected"
            : IsWaitingToReconnect ? "Waiting to reconnect..."
            : Engine.IsReconnecting ? "Reconnecting..."
            : IsIdle ? "Disconnected"
            : "Connecting...";
    }

    /// <summary>Which bot the Battle.net sign-in window is for, and which of its logins it saves to: several can ask at once.</summary>
    private string SignInWindowTitle()
    {
        var login = BattlenetCredentialProfileStore.Find(Config.BattlenetCredentialProfileId)?.DisplayLabel;
        return string.IsNullOrEmpty(login) || login == Config.DisplayName
            ? $"Battle.net Sign-In \u2014 {Config.DisplayName}"
            : $"Battle.net Sign-In \u2014 {Config.DisplayName} ({login})";
    }

    /// <summary>Swaps in an edited config (from the config window's Save) and refreshes anything derived from it, like the tab title.</summary>
    public void ApplyConfig(BotConfig newConfig)
    {
        // The edit started from a copy. A Battle.net profile the engine made for this bot meanwhile
        // (its first connect) isn't in it, and the picker has no way to choose "none", so an empty
        // one here only means the copy is older. Dropping it would sign the bot in from scratch.
        if (string.IsNullOrEmpty(newConfig.BattlenetCredentialProfileId))
        {
            newConfig.BattlenetCredentialProfileId = Engine.Config.BattlenetCredentialProfileId;
        }

        Engine.Config = newConfig;
        OnPropertyChanged(nameof(Config));
        OnPropertyChanged(nameof(Title));
        OnPropertyChanged(nameof(TabIconImage));
        OnPropertyChanged(nameof(HighlightBrush));
        OnPropertyChanged(nameof(BackgroundBrush));
        OnPropertyChanged(nameof(SupportsFriends));
        OnPropertyChanged(nameof(SupportsClan));
        OnPropertyChanged(nameof(IsSingleChannel));
        OnPropertyChanged(nameof(IsSc2));
        OnPropertyChanged(nameof(UserListNameFontSize));
        OnPropertyChanged(nameof(CanManageFriends));
        OnPropertyChanged(nameof(PortraitSize));
        OnPropertyChanged(nameof(ShowsMemberDetail));
        RefreshTheme();
    }

    public BotTabViewModel(BotEngine engine)
    {
        Engine = engine;
        ChatLineTrimmer.Attach(ChatLines, () => ChatLinesTrimmed?.Invoke());
        ShowStartupBanner();
        ChannelUsers.CollectionChanged += (_, _) => OnPropertyChanged(nameof(UsersTabHeader));
        Engine.Log += OnLog;
        Engine.SelfChatSent += OnSelfChatSent;
        Engine.ChatMessage += OnChatMessage;
        Engine.FriendsListUpdated += OnFriendsListUpdated;
        Engine.FriendInvitationsUpdated += OnFriendInvitationsUpdated;
        Engine.BncsConnected += () => Dispatcher.UIThread.Post(() =>
        {
            IsConnected = true;
            RefreshConnectionState();
            _ = OfferSc2PortraitsAsync();
        });
        Sc2PortraitStore.Downloaded += () => Dispatcher.UIThread.Post(() => IconVersion++);
        NativeMemberPortraits.Changed += OnPortraitsChanged;
        Engine.ActivityChanged += () => Dispatcher.UIThread.Post(RefreshConnectionState);
        Engine.DebugModeChanged += () => Dispatcher.UIThread.Post(() => OnPropertyChanged(nameof(DebugMode)));
        Engine.BncsDisconnected += _ => Dispatcher.UIThread.Post(() =>
        {
            MarkDisconnected();
            RefreshConnectionState();
        });
        // The Battle.net sign-in for SC2/SC:R/WC3:R, over the main window rather than this bot's
        // tab: it used to be set by the tab's view, so a bot whose tab wasn't showing (connected at
        // startup, reconnecting in the background) had no one to ask.
        Engine.Sc2ChallengeHandler = (url, token) => Sc2LoginChallenge.ShowAsync(url, SignInWindowTitle(), token);
        Engine.Sc2ChannelJoined += OnSc2ChannelJoined;
        Engine.Sc2ChannelLeft += OnSc2ChannelLeft;
        Engine.Sc2ChannelJoinRejected += OnSc2ChannelJoinRejected;
        Engine.Sc2ChannelActionFailed += OnSc2ChannelActionFailed;
        Engine.Sc2PublicChannelsReceived += OnSc2PublicChannelsReceived;
        IconOverrideStore.OverridesChanged += OnIconOverrideChanged;
        Invigoration.Core.Clan.ClanRosterStore.RosterChanged += OnClanRosterChanged;
        ThemeLibrary.ThemesChanged += OnThemesChanged;
        D2EquipmentStore.Changed += OnD2EquipmentChanged;
        RefreshClanRoster();
        RefreshTheme();
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
        IconVersion++;
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
    /// of parentheses plus "Invigoration STABLE Bunny" in red/green - shown
    /// once when a bot tab opens. It read "Beta bunny" through the whole
    /// 2.0.x line; the beta is over as of 2.1.0, and the bunny was promoted
    /// rather than retired. Ported from frmMain.frm's AddChat calls; the
    /// colored words reuse the same inline color-code marker (U+00A0 +
    /// letter) ChatColorFormatter already parses everywhere else.
    /// </summary>
    private void ShowStartupBanner()
    {
        var p = Engine.Palette;
        const string separator = "---------------------------------------------------";
        const char marker = ' ';
        var bunnyLine = $"Invigoration {marker}rSTABLE {marker}gBunny";

        ChatLines.Add(new ChatLineViewModel(separator, p.Highlight));
        ChatLines.Add(new ChatLineViewModel("()()", p.Info));
        ChatLines.Add(new ChatLineViewModel("(--)", p.Info));
        ChatLines.Add(new ChatLineViewModel("(')(')", p.Info));
        ChatLines.Add(new ChatLineViewModel(ChatColorFormatter.Parse(bunnyLine, p.Channel, p)));
        ChatLines.Add(new ChatLineViewModel($"C#/.NET port -- v{AppVersion.Current}", p.Debug));
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
        }

        RefreshConnectionState();
    }

    [RelayCommand]
    private Task DisconnectAsync() => Engine.DisconnectAsync();

    /// <summary>
    /// Diablo II's chat gem — the little jewel set in a socket beside the Send button, shown when
    /// the bot's theme turns it on (ChatTheme.ShowChatGem; the built-in Diablo II theme does). Blue when activated, red when not. Purely local flavor: it
    /// toggles its own color and prints one line into this bot's own chat log, and deliberately
    /// sends nothing to Battle.net, per explicit request.
    /// </summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ChatGemBrush))]
    public partial bool ChatGemActive { get; set; } = true;

    /// <summary>The gem's fill — the palette's own blue/red rather than hardcoded colors, so it follows whichever scheme the bot is using.</summary>
    public IBrush ChatGemBrush
    {
        get
        {
            var c = ChatGemActive ? Engine.Palette.Blue : Engine.Palette.Red;
            return new SolidColorBrush(Color.FromRgb(c.R, c.G, c.B));
        }
    }

    /// <summary>The running activation tally (BotConfig.ChatGemActivations), for the gem's tooltip.</summary>
    public string ChatGemTooltip =>
        $"Chat gem — click to toggle. Activated {Config.ChatGemActivations} time{(Config.ChatGemActivations == 1 ? "" : "s")} on this bot. Local decoration only; nothing is sent to Battle.net.";

    [RelayCommand]
    private void ToggleChatGem()
    {
        ChatGemActive = !ChatGemActive;
        var palette = Engine.Palette;
        var message = ChatGemActive ? "Chat Gem Activated" : "Chat Gem Deactivated";
        var color = ChatGemActive ? palette.Blue : palette.Red;
        (SelectedChannel?.ChatLines ?? ChatLines).Add(new ChatLineViewModel(message, color));

        if (ChatGemActive)
        {
            // Activations only — a toggle off isn't an activation. Persisted whenever SaveAll next
            // runs, same as every other BotConfig field (window close / Config window save).
            Config.ChatGemActivations++;
            OnPropertyChanged(nameof(ChatGemTooltip));

            // Separately, the opt-in leaderboard's own per-account monthly tally. Counted even
            // before (or without) consent so that agreeing partway through a month doesn't start
            // the user at zero — ChatGemTallyStore gates sharing, not counting, and nothing
            // transmits anywhere until an endpoint is configured. See ChatGemTallySender.
            if (!string.IsNullOrWhiteSpace(Config.Username))
            {
                ChatGemTallyStore.RecordActivation($"{Config.Username}@{Config.BattlenetServer}", DateTimeOffset.UtcNow);
            }
        }
    }

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
            SelectedChannel?.ChatLines.Add(new ChatLineViewModel(segments, ResolveSc2UserIcon("")));
        }
        else
        {
            ChatLines.Add(new ChatLineViewModel(segments, ResolveUserIcon(Config.Username)));
        }
    });

    private static bool IsUnreadWorthy(ChatEventType type) => type is ChatEventType.Talk or ChatEventType.Emote or ChatEventType.Broadcast;

    /// <summary>
    /// Events queued by OnChatMessage (called from whatever thread the engine's own network
    /// processing runs on), drained by ProcessPendingChatEvents on the UI thread. A
    /// ConcurrentQueue rather than a lock-protected List — enqueue and drain genuinely run on
    /// different threads concurrently, and this needs to stay cheap precisely during the
    /// scenario (a mass-join flood) it exists to help with.
    /// </summary>
    private readonly ConcurrentQueue<ChatEvent> _pendingChatEvents = new();

    /// <summary>0 = no drain currently scheduled/running, 1 = one is. CompareExchange-guarded so a flood of OnChatMessage calls schedules exactly one Dispatcher.Post, not one per event — see OnChatMessage's remarks for why that mattered.</summary>
    private int _chatEventDrainScheduled;

    /// <summary>
    /// Caps how many events one DrainPendingChatEvents call processes before yielding back to
    /// the Dispatcher (scheduling a follow-up drain for the rest) rather than draining the whole
    /// queue in one shot. Batching collapsed "2500 separate Dispatcher.Posts" down to "however
    /// many drains the flood needed" — but if all 2500 already landed in the queue before the
    /// very first drain got a chance to run (plausible: enqueueing from the network thread is
    /// much faster than the UI thread getting scheduled to look at the queue), "however many"
    /// would otherwise be exactly one, single 2500-event batch — back to one long UI-thread block,
    /// just with a different cause than the original per-event Dispatcher overhead. Chunking
    /// guarantees the window gets a real chance to repaint between chunks regardless of how large
    /// a burst arrives all at once.
    /// </summary>
    private const int MaxChatEventsPerDrain = 200;

    /// <summary>
    /// Queues the event and ensures exactly one drain is scheduled on the UI thread, rather than
    /// this class's previous approach of Dispatcher.UIThread.Post-ing individually per event.
    /// Each individual event's own processing was already fixed to be O(1) (the roster/channel-
    /// user-list indexing work elsewhere), but Avalonia's Dispatcher still has real per-Post
    /// scheduling/queuing overhead of its own — confirmed live as still a genuine source of a
    /// multi-second UI stall under a large mass-join burst (2500 joins), even with every
    /// individual event now cheap: 2500 separate queued Dispatcher operations is 2500 separate
    /// opportunities for the UI thread's own queue management overhead to add up, and every one
    /// of them sits ahead of the "please repaint the window" work in that same queue, which is
    /// what actually produces a beachball — the OS sees an unresponsive main thread, not slow
    /// C# code. Draining a whole batch in one Dispatcher operation collapses that back down to
    /// however many batches the flood actually needed (typically a small, roughly constant
    /// number), not one per event.
    /// </summary>
    private void OnChatMessage(ChatEvent e)
    {
        _pendingChatEvents.Enqueue(e);
        if (Interlocked.CompareExchange(ref _chatEventDrainScheduled, 1, 0) == 0)
        {
            Dispatcher.UIThread.Post(DrainPendingChatEvents);
        }
    }

    private void DrainPendingChatEvents()
    {
        // Reset before draining (not after): an event that arrives while this method is running
        // must see the flag already clear and successfully schedule its own follow-up drain,
        // rather than finding a drain "already scheduled" (this one, already past the point of
        // ever looking at it) and silently going unprocessed until something else happens to
        // enqueue later. The cost of that safety is possibly one redundant extra Post if timing
        // lines up just wrong — never a dropped or stalled event.
        Interlocked.Exchange(ref _chatEventDrainScheduled, 0);

        // Bounded by MaxChatEventsPerDrain (see its own remarks), not just "whatever's here right
        // now" — a flood still actively arriving from the network thread, or one that entirely
        // landed before this drain got scheduled, could otherwise keep this one Dispatcher
        // operation running for as long as the whole queue takes, which is just as capable of
        // starving the window's repaint as 2500 separate Posts was. Whatever's left after the cap
        // gets its own follow-up drain below instead.
        var count = Math.Min(_pendingChatEvents.Count, MaxChatEventsPerDrain);
        for (var i = 0; i < count && _pendingChatEvents.TryDequeue(out var e); i++)
        {
            ProcessChatEvent(e);
        }

        if (!_pendingChatEvents.IsEmpty && Interlocked.CompareExchange(ref _chatEventDrainScheduled, 1, 0) == 0)
        {
            Dispatcher.UIThread.Post(DrainPendingChatEvents);
        }
    }

    private void ProcessChatEvent(ChatEvent e)
    {
        if (SupportsMultiChannel && e.ChannelIndex is { } channelIndex)
        {
            var channel = Channels.FirstOrDefault(c => c.ChannelIndex == channelIndex);
            channel?.HandleChatEvent(e, Engine.Palette, ResolveSc2UserIcon(e.Username), Config.ShowUserIconsInChat, IsSc2 && Config.FullChatPortraits, hideNameCodes: IsSc2);
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
    }

    private void OnSc2ChannelJoined(byte channelIndex, ChatChannel channel, ObservableCollection<Person> users) =>
        Dispatcher.UIThread.Post(() =>
        {
            // Never two tabs for one index: every later lookup-by-index (leave, active-channel
            // tracking) only finds the first, leaving the second stuck. The same join repeated is
            // ignored; a tab left over from an earlier session is replaced, since a reconnect
            // reuses channel numbers and the old tab's roster belongs to the old session.
            var tab = new ChannelTabViewModel(channelIndex, channel, users);
            if (Channels.FirstOrDefault(c => c.ChannelIndex == channelIndex) is { } existing)
            {
                if (ReferenceEquals(existing.Users, users))
                {
                    return;
                }

                tab.AttachChatLineTrimmer();
                var wasSelected = SelectedChannel == existing;
                Channels[Channels.IndexOf(existing)] = tab;
                if (wasSelected || SelectedChannel is null)
                {
                    SelectedChannel = tab;
                }

                return;
            }

            tab.AttachChatLineTrimmer();
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

    /// <summary>Unsubscribes UsersTabHeader's count-tracking from whichever channel is being left, mirrored by the subscribe half in OnSelectedChannelChanged below.</summary>
    partial void OnSelectedChannelChanging(ChannelTabViewModel? oldValue, ChannelTabViewModel? newValue)
    {
        if (oldValue is not null)
        {
            oldValue.Users.CollectionChanged -= OnSelectedChannelUsersChanged;
        }
    }

    partial void OnSelectedChannelChanged(ChannelTabViewModel? value)
    {
        if (value is not null)
        {
            Engine.SetActiveSc2Channel(value.ChannelIndex);
            value.HasUnread = false;
            value.Users.CollectionChanged += OnSelectedChannelUsersChanged;
        }

        OnPropertyChanged(nameof(UsersTabHeader));
    }

    private void OnSelectedChannelUsersChanged(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e) =>
        OnPropertyChanged(nameof(UsersTabHeader));

    partial void OnCurrentChannelNameChanged(string value)
    {
        OnPropertyChanged(nameof(UsersTabHeader));
        Theme.ChannelName = value;
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
        var thread = WhisperThreads.FirstOrDefault(t => string.Equals(t.Peer, peer, StringComparison.OrdinalIgnoreCase));
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

        var thread = WhisperThreads.FirstOrDefault(t => string.Equals(t.Peer, peer, StringComparison.OrdinalIgnoreCase));
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
    public async Task SendWhisperAsync(WhisperThreadViewModel thread)
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
            friend.RealName = entry.RealName;
            friend.ModernStyle = SupportsMultiChannel;
            friend.ShowRealName = Config.ShowFriendRealNames;
            friend.ShowOffline = Config.ShowOfflineFriends;
            friend.CanRemove = CanManageFriends;
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
                _channelUsersByName.Clear();
                CurrentChannelName = e.Text;
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
                if (_channelUsersByName.Remove(e.Username, out var leaving))
                {
                    ChannelUsers.Remove(leaving);
                }

                ChatLines.Add(new ChatLineViewModel($"*** {e.Username} has left the channel.", palette.Gray));
                break;

            case ChatEventType.UserFlags:
                UpsertUser(e);
                break;

            case ChatEventType.Talk:
                // Discord users show under their own name with the Discord logo: either the message
                // came in over this bot's own bridge (speaker "[Discord] name"), or another
                // Invigoration bot relayed it into the channel ("[Discord] name: message").
                if (e.Origin == ChatEventOrigin.Discord && DiscordRelayLine.DiscordUserFromSpeaker(e.Username) is { } fromOwnBridge)
                {
                    ChatLines.Add(DiscordRelayRendering.Build(fromOwnBridge, e.Text, relayedBy: null, palette, Config.ShowUserIconsInChat));
                }
                else if (DiscordRelayLine.TryParse(e.Text, out var relayedUser, out var relayedText))
                {
                    ChatLines.Add(DiscordRelayRendering.Build(relayedUser, relayedText, relayedBy: e.Username, palette, Config.ShowUserIconsInChat));
                }
                else
                {
                    ChatLines.Add(new ChatLineViewModel(BuildUserLine(e.Username, e.Text, e.Flags, palette), ResolveUserIcon(e.Username)));
                }

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
        if (_channelUsersByName.TryGetValue(e.Username, out var user))
        {
            ChannelUsers.Remove(user);
        }
        else
        {
            user = new ChannelUserViewModel(e.Username)
            {
                UseClassicIconStyle = Config.ClassicUserIconStyle,
                ShowLadderIcons = Config.ShowLadderIcons,
                BotProduct = Config.Product,
                Palette = Engine.Palette,
                UseD2Layout = Theme.UsesCharacterDock,
            };
            _channelUsersByName[e.Username] = user;
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

    /// <summary>Whether the Friends tab can add, remove and answer Battle.net friends: SC:R bots, so far (SC2's commands aren't mapped).</summary>
    public bool CanManageFriends => IsSingleChannel;

    /// <summary>Pending Battle.net friend requests, to this bot's account or sent from it.</summary>
    public ObservableCollection<FriendInvitationViewModel> FriendInvitations { get; } = [];

    /// <summary>The Friends tab's "Add friend" box.</summary>
    [ObservableProperty]
    public partial string NewFriendBattleTag { get; set; } = "";

    [RelayCommand]
    private void AddFriend()
    {
        if (Engine.AddBattlenetFriend(NewFriendBattleTag))
        {
            NewFriendBattleTag = "";
        }
    }

    [RelayCommand]
    private void AcceptFriendInvitation(FriendInvitationViewModel invitation) => Engine.AnswerFriendInvitation(invitation.Id, accept: true);

    [RelayCommand]
    private void DeclineFriendInvitation(FriendInvitationViewModel invitation) => Engine.AnswerFriendInvitation(invitation.Id, accept: false);

    public void RemoveFriend(FriendEntryViewModel friend) => Engine.RemoveBattlenetFriend(friend.Account);

    private void OnFriendInvitationsUpdated(IReadOnlyList<Invigoration.Core.Sc2.FriendInvitation> invitations) => Dispatcher.UIThread.Post(() =>
    {
        FriendInvitations.Clear();
        foreach (var invitation in invitations)
        {
            FriendInvitations.Add(new FriendInvitationViewModel(invitation.Id, invitation.BattleTag, invitation.Sent));
        }

        OnPropertyChanged(nameof(HasFriendInvitations));
    });

    public bool HasFriendInvitations => FriendInvitations.Count > 0;

    /// <summary>The Friends tab's "Real names" checkbox (Battle.net 2.0 friends only). Saved with the bot's config.</summary>
    public bool ShowFriendRealNames
    {
        get => Config.ShowFriendRealNames;
        set
        {
            if (Config.ShowFriendRealNames == value)
            {
                return;
            }

            Config.ShowFriendRealNames = value;
            foreach (var friend in Friends)
            {
                friend.ShowRealName = value;
            }

            OnPropertyChanged();
        }
    }

    /// <summary>The Friends tab's "Offline" checkbox. Saved with the bot's config.</summary>
    public bool ShowOfflineFriends
    {
        get => Config.ShowOfflineFriends;
        set
        {
            if (Config.ShowOfflineFriends == value)
            {
                return;
            }

            Config.ShowOfflineFriends = value;
            foreach (var friend in Friends)
            {
                friend.ShowOffline = value;
            }

            OnPropertyChanged();
        }
    }

    /// <summary>Flips Config.ShowLadderIcons and updates every row now; the right-click "Ladder Icons" item. Saved like ToggleClassicUserIconStyle.</summary>
    public void ToggleShowLadderIcons()
    {
        Config.ShowLadderIcons = !Config.ShowLadderIcons;
        foreach (var user in ChannelUsers)
        {
            user.ShowLadderIcons = Config.ShowLadderIcons;
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

        var statString = _channelUsersByName.GetValueOrDefault(username)?.StatString;
        if (string.IsNullOrEmpty(statString))
        {
            return null;
        }

        var key = ChatIcon.GetProductIconKey(statString);
        return string.IsNullOrEmpty(key) ? null : GameIconLoader.Get(key);
    }

    /// <summary>
    /// SC2/SC:R speaker icon (BotConfig.ShowUserIconsInChat): an SC:R speaker's classic game icon, an
    /// SC2 speaker's portrait once known, otherwise this bot's own game icon.
    /// </summary>
    private Bitmap? ResolveSc2UserIcon(string username)
    {
        if (!Config.ShowUserIconsInChat)
        {
            return null;
        }

        if (NativeMemberProducts.IconKeyFor(username) is { } key)
        {
            return GameIconLoader.Get(key);
        }

        if (NativeMemberPortraits.For(username) is { } portrait && Sc2PortraitImages.Get(portrait.Sheet, portrait.Cell) is { } image)
        {
            return image;
        }

        return GameIconLoader.Get(BncsProduct.GetIconKey(Config.Product));
    }

    public ValueTask DisposeAsync()
    {
        Engine.Log -= OnLog;
        Engine.SelfChatSent -= OnSelfChatSent;
        Engine.ChatMessage -= OnChatMessage;
        Engine.FriendsListUpdated -= OnFriendsListUpdated;
        Engine.FriendInvitationsUpdated -= OnFriendInvitationsUpdated;
        Engine.Sc2ChannelJoined -= OnSc2ChannelJoined;
        Engine.Sc2ChannelLeft -= OnSc2ChannelLeft;
        Engine.Sc2ChannelJoinRejected -= OnSc2ChannelJoinRejected;
        Engine.Sc2ChannelActionFailed -= OnSc2ChannelActionFailed;
        Engine.Sc2PublicChannelsReceived -= OnSc2PublicChannelsReceived;
        IconOverrideStore.OverridesChanged -= OnIconOverrideChanged;
        Invigoration.Core.Clan.ClanRosterStore.RosterChanged -= OnClanRosterChanged;
        ThemeLibrary.ThemesChanged -= OnThemesChanged;
        D2EquipmentStore.Changed -= OnD2EquipmentChanged;
        return Engine.DisposeAsync();
    }
}
