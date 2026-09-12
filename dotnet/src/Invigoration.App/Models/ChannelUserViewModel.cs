using Avalonia.Media;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.Core.Chat;

namespace Invigoration.App.Models;

public partial class ChannelUserViewModel(string username) : ObservableObject
{
    public string Username { get; } = username;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(StatusIconImage))]
    [NotifyPropertyChangedFor(nameof(ProductIconImage))]
    [NotifyPropertyChangedFor(nameof(DisplayIconImage))]
    [NotifyPropertyChangedFor(nameof(ShowSeparateStatusIcon))]
    [NotifyPropertyChangedFor(nameof(IsLargeIcon))]
    [NotifyPropertyChangedFor(nameof(ShowLadderScore))]
    [NotifyPropertyChangedFor(nameof(UsernameBrush))]
    public partial uint Flags { get; set; }

    [ObservableProperty]
    public partial int Ping { get; set; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ProductIconImage))]
    [NotifyPropertyChangedFor(nameof(DisplayIconImage))]
    [NotifyPropertyChangedFor(nameof(IsLargeIcon))]
    [NotifyPropertyChangedFor(nameof(LadderScoreText))]
    [NotifyPropertyChangedFor(nameof(ShowLadderScore))]
    [NotifyPropertyChangedFor(nameof(UsernameBrush))]
    public partial string StatString { get; set; } = "";

    /// <summary>The local bot's own 4-char wire-order product code (BotConfig.Product) — pushed in by BotTabViewModel.UpsertUser at row construction, same reason as UseClassicIconStyle below (this row's DataContext has no reachable path back up to the tab). Used only to decide whether this row counts as "same game as you" for UsernameBrush.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(UsernameBrush))]
    public partial string BotProduct { get; set; } = "";

    /// <summary>The bot's active chat color scheme (Engine.Palette) — pushed in alongside BotProduct, same reason.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(UsernameBrush))]
    public partial ChatPalette Palette { get; set; } = ChatPalette.Invigoration;

    /// <summary>Mirrors BotConfig.ClassicUserIconStyle — pushed in by BotTabViewModel (see UpsertUser/ApplyClassicUserIconStyle) rather than read from Config directly, since this row's own DataContext has no reachable path back up to it (the same ancestor-binding pitfall noted throughout BotTabView.axaml).</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(DisplayIconImage))]
    [NotifyPropertyChangedFor(nameof(ShowSeparateStatusIcon))]
    [NotifyPropertyChangedFor(nameof(IsLargeIcon))]
    public partial bool UseClassicIconStyle { get; set; }

    public Bitmap? ProductIconImage => GameIconLoader.Get(ChatIcon.GetProductIconKey(StatString, Flags));

    public Bitmap? StatusIconImage => GameIconLoader.Get(ChatIcon.GetStatusIconKey(Flags));

    /// <summary>The game-icon slot's actual image: the classic Battle.net behavior of a rank badge replacing the game icon entirely when UseClassicIconStyle is on and one applies, otherwise always the game icon (with the badge shown separately — see ShowSeparateStatusIcon).</summary>
    public Bitmap? DisplayIconImage => UseClassicIconStyle && StatusIconImage is not null ? StatusIconImage : ProductIconImage;

    /// <summary>Whether the row's separate status-badge slot should still show — same "is there actually a badge" check as before classic style existed, plus: never when classic style already folded the badge into DisplayIconImage instead.</summary>
    public bool ShowSeparateStatusIcon => StatusIconImage is not null && !UseClassicIconStyle;

    /// <summary>True once a bigger-than-classic icon (e.g. a 64x64 override, or a 64x64 status badge in classic icon style) is in play, so the row template can switch from one inline line to username/ping stacked — the tall icon otherwise dwarfs a single text line. Checks whatever's actually displayed (DisplayIconImage), not always ProductIconImage, since classic style can swap in a differently-sized badge.</summary>
    public bool IsLargeIcon => DisplayIconImage is { PixelSize.Height: > 16 };

    /// <summary>The ladder rating to stamp as text over the product-icon slot (real classic Battle.net behavior — see ChatIcon.GetLadderScore), or "" when there's nothing to show.</summary>
    public string LadderScoreText => ChatIcon.GetLadderScore(StatString)?.ToString() ?? "";

    /// <summary>Only when there's a score AND the product-icon slot is actually showing the win/rank plate it belongs on — StatusIconImage not null means GetProductIconKey already deferred to the flat generic logo instead (same status-icon-takes-priority rule it applies internally), so stamping a score on that would be wrong.</summary>
    public bool ShowLadderScore => LadderScoreText != "" && StatusIconImage is null;

    private bool IsSameProduct => StatString.Length >= 4 && BotProduct.Length >= 4 &&
                                   StatString.AsSpan(0, 4).SequenceEqual(BotProduct.AsSpan(0, 4));

    /// <summary>Real classic Battle.net channel-list username coloring — see ChatPalette.GetChannelListNameColor's own remarks for the exact rule (Blizzard rep/Admin/same-game/everyone-else).</summary>
    public IBrush UsernameBrush
    {
        get
        {
            var c = Palette.GetChannelListNameColor(Flags, IsSameProduct);
            return new SolidColorBrush(Color.FromRgb(c.R, c.G, c.B));
        }
    }

    /// <summary>Re-raises change notifications for the icon-derived properties — called after an override is applied/reset so an already-populated user list updates immediately instead of needing a reconnect.</summary>
    public void RefreshIcons()
    {
        OnPropertyChanged(nameof(ProductIconImage));
        OnPropertyChanged(nameof(StatusIconImage));
        OnPropertyChanged(nameof(DisplayIconImage));
        OnPropertyChanged(nameof(ShowSeparateStatusIcon));
        OnPropertyChanged(nameof(IsLargeIcon));
        OnPropertyChanged(nameof(LadderScoreText));
        OnPropertyChanged(nameof(ShowLadderScore));
    }
}
