using System.Collections.ObjectModel;
using Avalonia.Media;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.Models;
using Invigoration.Core;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;

namespace Invigoration.App.ViewModels;

/// <summary>One bot tab: wraps a BotEngine and projects its events onto observable collections for binding.</summary>
public partial class BotTabViewModel : ViewModelBase, IAsyncDisposable
{
    public BotEngine Engine { get; }

    public BotConfig Config => Engine.Config;

    public string Title => Config.DisplayName;

    /// <summary>The active bot's scheme-specific accent, for marking this tab as the open one and/or the chat input as focused.</summary>
    public IBrush HighlightBrush => new SolidColorBrush(
        Color.FromRgb(Engine.Palette.Highlight.R, Engine.Palette.Highlight.G, Engine.Palette.Highlight.B));

    /// <summary>The active bot's chat-log background, from its selected color scheme.</summary>
    public IBrush BackgroundBrush => new SolidColorBrush(
        Color.FromRgb(Engine.Palette.Background.R, Engine.Palette.Background.G, Engine.Palette.Background.B));

    public ObservableCollection<ChatLineViewModel> ChatLines { get; } = [];

    public ObservableCollection<ChannelUserViewModel> ChannelUsers { get; } = [];

    public ObservableCollection<FriendEntryViewModel> Friends { get; } = [];

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
    });

    /// <summary>
    /// Restores the VB6 original's Form_Load ASCII-art banner - a bunny made
    /// of parentheses plus "Invigoration Nightly Bunny" in red/green - shown
    /// once when a bot tab opens, as a nod to this project's long-running
    /// beta status. Ported from frmMain.frm's AddChat calls; the colored
    /// "Nightly"/"Bunny" words reuse the same inline color-code marker
    /// (U+00A0 + letter) ChatColorFormatter already parses everywhere else.
    /// </summary>
    private void ShowStartupBanner()
    {
        var p = Engine.Palette;
        const string separator = "---------------------------------------------------";
        const char marker = ' ';
        var bunnyLine = $"Invigoration {marker}rNightly {marker}gBunny";

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

        var triggerChar = Config.Trigger.FirstOrDefault();
        var isTriggered = text.Length > 0 && (text[0] == triggerChar || text[0] == '/');

        if (isTriggered && text.Length > 1 && text[0] == text[1])
        {
            // Doubling the leading trigger/slash escapes it: sends the rest
            // verbatim as a real chat message instead of intercepting it as a
            // local-only command. Lets you test another bot's commands (e.g.
            // "!!trivia join") from this bot's own tab as if it were just
            // another channel member, instead of it silently running against
            // this bot's own (likely idle) engine.
            await Engine.SendChatCommandAsync(text[1..]);
        }
        else if (isTriggered)
        {
            await Engine.RunLocalCommandAsync(text);
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

    private void OnChatMessage(ChatEvent e) => Dispatcher.UIThread.Post(() => HandleChatEvent(e));

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
    });

    private void HandleChatEvent(ChatEvent e)
    {
        var palette = Engine.Palette;
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
                ChatLines.Add(new ChatLineViewModel(BuildUserLine(e.Username, e.Text, e.Flags, palette)));
                break;

            case ChatEventType.Emote:
                ChatLines.Add(new ChatLineViewModel($"<{e.Username} {e.Text}>", palette.GetEmoteColor(e.Flags)));
                break;

            case ChatEventType.Whisper:
                ChatLines.Add(new ChatLineViewModel($"[{e.Username} whispers]: {e.Text}", palette.Whisper));
                break;

            case ChatEventType.WhisperSent:
                ChatLines.Add(new ChatLineViewModel($"[whisper to {e.Username}]: {e.Text}", palette.Whisper));
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

    private void UpsertUser(ChatEvent e)
    {
        var user = ChannelUsers.FirstOrDefault(u => u.Username == e.Username);
        if (user is null)
        {
            user = new ChannelUserViewModel(e.Username);
            ChannelUsers.Add(user);
        }

        user.Flags = e.Flags;
        user.Ping = e.Ping;
        if (e.Text.Length > 0)
        {
            user.StatString = e.Text;
        }
    }

    private static IReadOnlyList<ChatLogSegment> BuildUserLine(string username, string text, uint flags, ChatPalette palette)
    {
        var segments = new List<ChatLogSegment> { new(palette.GetUserNameColor(flags), $"{username}: ") };
        segments.AddRange(ChatColorFormatter.Parse(text, palette.GetChatColor(flags), palette));
        return segments;
    }

    public ValueTask DisposeAsync()
    {
        Engine.Log -= OnLog;
        Engine.ChatMessage -= OnChatMessage;
        Engine.FriendsListUpdated -= OnFriendsListUpdated;
        IconOverrideStore.OverridesChanged -= OnIconOverrideChanged;
        Invigoration.Core.Clan.ClanRosterStore.RosterChanged -= OnClanRosterChanged;
        return Engine.DisposeAsync();
    }
}
