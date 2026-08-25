using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.Core.Chat;

namespace Invigoration.App.Models;

public partial class ChannelUserViewModel(string username) : ObservableObject
{
    public string Username { get; } = username;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(StatusIconImage))]
    [NotifyPropertyChangedFor(nameof(DisplayIconImage))]
    [NotifyPropertyChangedFor(nameof(ShowSeparateStatusIcon))]
    [NotifyPropertyChangedFor(nameof(IsLargeIcon))]
    public partial uint Flags { get; set; }

    [ObservableProperty]
    public partial int Ping { get; set; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ProductIconImage))]
    [NotifyPropertyChangedFor(nameof(DisplayIconImage))]
    [NotifyPropertyChangedFor(nameof(IsLargeIcon))]
    public partial string StatString { get; set; } = "";

    /// <summary>Mirrors BotConfig.ClassicUserIconStyle — pushed in by BotTabViewModel (see UpsertUser/ApplyClassicUserIconStyle) rather than read from Config directly, since this row's own DataContext has no reachable path back up to it (the same ancestor-binding pitfall noted throughout BotTabView.axaml).</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(DisplayIconImage))]
    [NotifyPropertyChangedFor(nameof(ShowSeparateStatusIcon))]
    [NotifyPropertyChangedFor(nameof(IsLargeIcon))]
    public partial bool UseClassicIconStyle { get; set; }

    public Bitmap? ProductIconImage => GameIconLoader.Get(ChatIcon.GetProductIconKey(StatString));

    public Bitmap? StatusIconImage => GameIconLoader.Get(ChatIcon.GetStatusIconKey(Flags));

    /// <summary>The game-icon slot's actual image: the classic Battle.net behavior of a rank badge replacing the game icon entirely when UseClassicIconStyle is on and one applies, otherwise always the game icon (with the badge shown separately — see ShowSeparateStatusIcon).</summary>
    public Bitmap? DisplayIconImage => UseClassicIconStyle && StatusIconImage is not null ? StatusIconImage : ProductIconImage;

    /// <summary>Whether the row's separate status-badge slot should still show — same "is there actually a badge" check as before classic style existed, plus: never when classic style already folded the badge into DisplayIconImage instead.</summary>
    public bool ShowSeparateStatusIcon => StatusIconImage is not null && !UseClassicIconStyle;

    /// <summary>True once a bigger-than-classic icon (e.g. a 64x64 override, or a 64x64 status badge in classic icon style) is in play, so the row template can switch from one inline line to username/ping stacked — the tall icon otherwise dwarfs a single text line. Checks whatever's actually displayed (DisplayIconImage), not always ProductIconImage, since classic style can swap in a differently-sized badge.</summary>
    public bool IsLargeIcon => DisplayIconImage is { PixelSize.Height: > 16 };

    /// <summary>Re-raises change notifications for the icon-derived properties — called after an override is applied/reset so an already-populated user list updates immediately instead of needing a reconnect.</summary>
    public void RefreshIcons()
    {
        OnPropertyChanged(nameof(ProductIconImage));
        OnPropertyChanged(nameof(StatusIconImage));
        OnPropertyChanged(nameof(DisplayIconImage));
        OnPropertyChanged(nameof(ShowSeparateStatusIcon));
        OnPropertyChanged(nameof(IsLargeIcon));
    }
}
