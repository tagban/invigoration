using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.Core.Chat;

namespace Invigoration.App.Models;

public partial class FriendEntryViewModel(string account) : ObservableObject
{
    public string Account { get; } = account;

    [ObservableProperty]
    public partial FriendStatus Status { get; set; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsOnline))]
    [NotifyPropertyChangedFor(nameof(ProductIconImage))]
    [NotifyPropertyChangedFor(nameof(StatusText))]
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
