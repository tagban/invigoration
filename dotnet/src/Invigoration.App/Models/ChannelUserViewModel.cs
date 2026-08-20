using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.Core.Chat;

namespace Invigoration.App.Models;

public partial class ChannelUserViewModel(string username) : ObservableObject
{
    public string Username { get; } = username;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(StatusIconImage))]
    public partial uint Flags { get; set; }

    [ObservableProperty]
    public partial int Ping { get; set; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ProductIconImage))]
    [NotifyPropertyChangedFor(nameof(IsLargeIcon))]
    public partial string StatString { get; set; } = "";

    public Bitmap? ProductIconImage => GameIconLoader.Get(ChatIcon.GetProductIconKey(StatString));

    public Bitmap? StatusIconImage => GameIconLoader.Get(ChatIcon.GetStatusIconKey(Flags));

    /// <summary>True once a bigger-than-classic icon (e.g. a 64x64 override) is in play, so the row template can switch from one inline line to username/ping stacked — the tall icon otherwise dwarfs a single text line.</summary>
    public bool IsLargeIcon => ProductIconImage is { PixelSize.Height: > 16 };

    /// <summary>Re-raises change notifications for the icon-derived properties — called after an override is applied/reset so an already-populated user list updates immediately instead of needing a reconnect.</summary>
    public void RefreshIcons()
    {
        OnPropertyChanged(nameof(ProductIconImage));
        OnPropertyChanged(nameof(StatusIconImage));
        OnPropertyChanged(nameof(IsLargeIcon));
    }
}
