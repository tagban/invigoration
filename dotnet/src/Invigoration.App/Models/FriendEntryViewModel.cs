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

    public bool IsOnline => Location != FriendLocation.Offline;

    /// <summary>The product icon while online; a simple gray "offline" indicator (its own overridable key) once they've gone offline, since there's no product code to show an icon for at that point.</summary>
    public Bitmap? ProductIconImage => IsOnline
        ? GameIconLoader.Get(ChatIcon.GetProductIconKey(ProductCode))
        : GameIconLoader.Get("offline");

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

    public string StatusText => Location switch
    {
        FriendLocation.Offline => "Offline",
        FriendLocation.NotInChat => "Online",
        FriendLocation.InChat => string.IsNullOrEmpty(LocationName) ? "In chat" : $"In channel: {LocationName}",
        FriendLocation.PublicGame => string.IsNullOrEmpty(LocationName) ? "In a public game" : $"Playing: {LocationName}",
        FriendLocation.PrivateGame or FriendLocation.PrivateGameMutual => "In a private game",
        _ => "",
    };
}
