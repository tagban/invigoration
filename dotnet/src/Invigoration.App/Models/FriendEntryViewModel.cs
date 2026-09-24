using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.Core.Chat;

namespace Invigoration.App.Models;

public partial class FriendEntryViewModel(string account) : ObservableObject
{
    public string Account { get; } = account;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(PresenceState))]
    public partial FriendStatus Status { get; set; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsOnline))]
    [NotifyPropertyChangedFor(nameof(IsListed))]
    [NotifyPropertyChangedFor(nameof(ProductIconImage))]
    [NotifyPropertyChangedFor(nameof(StatusText))]
    [NotifyPropertyChangedFor(nameof(PresenceState))]
    public partial FriendLocation Location { get; set; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ProductIconImage))]
    public partial string ProductCode { get; set; } = "";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(StatusText))]
    public partial string LocationName { get; set; } = "";

    /// <summary>Shown after the name, Battle.net-app style, when a Battle.net 2.0 friend shares it.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(DisplayName))]
    [NotifyPropertyChangedFor(nameof(NamePart))]
    [NotifyPropertyChangedFor(nameof(CodePart))]
    public partial string RealName { get; set; } = "";

    /// <summary>
    /// A Battle.net 2.0 friend (SC2/SC:R bots): shown like the Battle.net app, as "BattleTag (Real
    /// Name)" over the game they're in, with no classic presence dot.
    /// </summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(StatusText))]
    [NotifyPropertyChangedFor(nameof(ShowsPresenceDot))]
    [NotifyPropertyChangedFor(nameof(ProductIconImage))]
    [NotifyPropertyChangedFor(nameof(NameFontSize))]
    public partial bool ModernStyle { get; set; }

    /// <summary>Show friends by their real name, when they share it, instead of their BattleTag (BotConfig.ShowFriendRealNames; off by default).</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(DisplayName))]
    [NotifyPropertyChangedFor(nameof(NamePart))]
    [NotifyPropertyChangedFor(nameof(CodePart))]
    public partial bool ShowRealName { get; set; }

    /// <summary>List offline friends too (BotConfig.ShowOfflineFriends; off by default).</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsListed))]
    public partial bool ShowOffline { get; set; }

    public bool IsListed => IsOnline || ShowOffline;

    /// <summary>Whether the bot can remove this friend (a Battle.net friend on an SC:R bot).</summary>
    [ObservableProperty]
    public partial bool CanRemove { get; set; }

    /// <summary>"Remove Friend" was clicked once; the popup now asks to confirm.</summary>
    [ObservableProperty]
    public partial bool ConfirmingRemove { get; set; }

    private bool UsesRealName => ShowRealName && RealName.Length > 0;

    public string DisplayName => UsesRealName ? RealName : Account;

    /// <summary>The name, without a BattleTag's "#1234" code, which <see cref="CodePart"/> shows dimmed.</summary>
    public string NamePart => UsesRealName ? RealName : Core.Chat.NameParts.Split(Account).Name;

    public string CodePart => UsesRealName ? "" : Core.Chat.NameParts.Split(Account).Code;

    /// <summary>Kept for the row layout; the real name now replaces the BattleTag rather than following it.</summary>
    public string RealNamePart => "";

    public double NameFontSize => ModernStyle ? 12 : 14;

    public bool ShowsPresenceDot => !ModernStyle;

    public bool IsOnline => Location != FriendLocation.Offline;

    /// <summary>
    /// The product icon while online; a simple gray "offline" indicator (its own overridable key) once
    /// they've gone offline. A Battle.net 2.0 friend's code may be an icon key itself ("d4", "bnet2").
    /// </summary>
    /// <remarks>Battle.net 2.0 friends always use the Battle.net 2.0 icon set: they come from Battle.net's friends list, whatever set the bot shows.</remarks>
    public Bitmap? ProductIconImage
    {
        get
        {
            var key = !IsOnline ? "offline" : ChatIcon.GetProductIconKey(ProductCode) is { Length: > 0 } product ? product : ProductCode;
            return ModernStyle ? IconSets.GetBnet2(key) : GameIconLoader.Get(key);
        }
    }

    /// <summary>Re-raises the icon-derived property so an already-populated friends list updates immediately after an override is applied/reset, without needing a reconnect.</summary>
    public void RefreshIcon() => OnPropertyChanged(nameof(ProductIconImage));

    /// <summary>Bound to the right-click "Whisper" inline compose popup's textbox — see BotTabView.axaml's ContextFlyout on the Friends row.</summary>
    [ObservableProperty]
    public partial string WhisperDraft { get; set; } = "";

    /// <summary>Old-school Battle.net presence, derived from the same Status/Location flags StatusText reads — DoNotDisturb/Away take priority over a plain "in chat" location, matching how classic Battle.net clients rendered these as distinct status icons rather than just text.</summary>
    public PresenceState PresenceState => Location switch
    {
        FriendLocation.Offline => PresenceState.Offline,
        FriendLocation.PublicGame or FriendLocation.PrivateGame or FriendLocation.PrivateGameMutual => PresenceState.InGame,
        _ when Status.HasFlag(FriendStatus.DoNotDisturb) => PresenceState.DoNotDisturb,
        _ when Status.HasFlag(FriendStatus.Away) => PresenceState.Away,
        _ => PresenceState.Available,
    };

    public string StatusText => ModernStyle
        ? (IsOnline ? (LocationName.Length > 0 ? LocationName : "Online") : "Offline")
        : Location switch
    {
        FriendLocation.Offline => "Offline",
        FriendLocation.NotInChat => "Online",
        FriendLocation.InChat => string.IsNullOrEmpty(LocationName) ? "In chat" : $"In channel: {LocationName}",
        FriendLocation.PublicGame => string.IsNullOrEmpty(LocationName) ? "In a public game" : $"Playing: {LocationName}",
        FriendLocation.PrivateGame or FriendLocation.PrivateGameMutual => "In a private game",
        _ => "",
    };
}
