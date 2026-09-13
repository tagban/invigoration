using Avalonia;
using Avalonia.Layout;
using Avalonia.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.Core.Config;

namespace Invigoration.App.ViewModels;

/// <summary>Anything that shows a themed chat surface — a bot's tab, or the theme manager's preview — so the shared input-row styles (ThemeStyles.axaml) can bind to its theme without caring which.</summary>
public interface IThemedSurface
{
    ChatThemeViewModel Theme { get; }
}

/// <summary>
/// A <see cref="ChatTheme"/> turned into what the views draw with: brushes built from its seven
/// material colors, and every layout coordinate that depends on its frame and layout, so
/// BotTabView binds straight to values instead of combining booleans through converters.
/// Immutable apart from <see cref="ChannelName"/> — a changed theme is a new instance.
/// </summary>
public sealed partial class ChatThemeViewModel : ObservableObject
{
    // The frame's own gutter around the chat and input: every frame design draws inside these
    // bands, so none of them can ever paint over text. The top is taller for the name plate.
    private static readonly Thickness FramedChatMargin = new(18, 26, 18, 4);
    private static readonly Thickness FramedChatMarginNoPlate = new(18, 18, 18, 4);
    private static readonly Thickness FramedInputMargin = new(18, 17, 18, 18);

    public ChatThemeViewModel(ChatTheme theme)
    {
        Theme = theme;
        var m = theme.Materials;
        var band = ToColor(m.Band);
        var bandShade = ToColor(m.BandShade);
        var trim = ToColor(m.Trim);
        var trimShade = ToColor(m.TrimShade);
        var metal = ToColor(m.Metal);
        var accent = ToColor(m.Accent);

        BandBrush = Gradient(new RelativePoint(0, 0, RelativeUnit.Relative), new RelativePoint(1, 1, RelativeUnit.Relative),
            band, Mix(band, bandShade, 0.55), bandShade);
        CapBrush = Gradient(new RelativePoint(0, 0, RelativeUnit.Relative), new RelativePoint(1, 1, RelativeUnit.Relative),
            Mix(band, Colors.White, 0.22), Mix(band, bandShade, 0.5), Mix(bandShade, Colors.Black, 0.15));
        TrimBrush = Gradient(new RelativePoint(0, 0, RelativeUnit.Relative), new RelativePoint(0, 1, RelativeUnit.Relative),
            Mix(trim, Colors.White, 0.35), trim, trimShade);
        TrimSolidBrush = new SolidColorBrush(trim);
        TrimShadeSolidBrush = new SolidColorBrush(trimShade);
        BevelBrush = new SolidColorBrush(Mix(band, Colors.White, 0.28));
        MetalPlateBrush = Gradient(new RelativePoint(0, 0, RelativeUnit.Relative), new RelativePoint(0, 1, RelativeUnit.Relative),
            Mix(metal, Colors.White, 0.3), metal, Mix(metal, Colors.Black, 0.5));
        StudBrush = new RadialGradientBrush
        {
            GradientOrigin = new RelativePoint(0.3, 0.3, RelativeUnit.Relative),
            Center = new RelativePoint(0.4, 0.4, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(Mix(metal, Colors.White, 0.6), 0),
                new GradientStop(metal, 0.55),
                new GradientStop(Mix(metal, Colors.Black, 0.6), 1),
            },
        };
        AccentBrush = new SolidColorBrush(accent);
        AccentGlowBrush = new SolidColorBrush(Color.FromArgb(0xA0, accent.R, accent.G, accent.B));
        WellBrush = new SolidColorBrush(ToColor(m.Well));
        HeaderFontFamily = string.IsNullOrWhiteSpace(theme.HeaderFont) ? FontFamily.Default : new FontFamily(theme.HeaderFont);
    }

    public ChatTheme Theme { get; }

    /// <summary>The channel on the name plate — pushed in by whoever shows this theme, since the theme itself doesn't know.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ShowNamePlate))]
    public partial string ChannelName { get; set; } = "";

    public bool HasFrame => Theme.Frame != ThemeFrameStyle.None;

    public bool IsDiabloFrame => Theme.Frame == ThemeFrameStyle.DiabloII;

    public bool IsStarCraftFrame => Theme.Frame == ThemeFrameStyle.StarCraft;

    public bool IsWarcraftFrame => Theme.Frame == ThemeFrameStyle.Warcraft;

    public bool UsesCharacterDock => Theme.Layout == ThemeLayout.CharacterDock;

    public bool ShowChatGem => Theme.ShowChatGem;

    public bool ShowNamePlate => HasFrame && Theme.ShowNamePlate && !string.IsNullOrEmpty(ChannelName);

    // --- Brushes ---

    public IBrush BandBrush { get; }

    public IBrush CapBrush { get; }

    public IBrush TrimBrush { get; }

    public IBrush TrimSolidBrush { get; }

    public IBrush TrimShadeSolidBrush { get; }

    public IBrush BevelBrush { get; }

    public IBrush MetalPlateBrush { get; }

    public IBrush StudBrush { get; }

    public IBrush AccentBrush { get; }

    public IBrush AccentGlowBrush { get; }

    public IBrush WellBrush { get; }

    public FontFamily HeaderFontFamily { get; }

    // --- Layout (see BotTabView.axaml's root grid remarks) ---

    /// <summary>The chat spans the full width over a character dock, otherwise only the left column.</summary>
    public int ChatColumnSpan => UsesCharacterDock ? 3 : 1;

    public Thickness ChatMargin => !HasFrame ? default : Theme.ShowNamePlate ? FramedChatMargin : FramedChatMarginNoPlate;

    /// <summary>The input sits directly under the chat (row 2) whenever a frame has to wrap both, or a dock takes the bottom row; otherwise in its usual bottom row.</summary>
    public int InputRow => HasFrame || UsesCharacterDock ? 2 : 3;

    /// <summary>A framed standard layout keeps the input under the chat column only, so one rectangular frame can wrap both while the user list stays alongside.</summary>
    public int InputColumnSpan => HasFrame && !UsesCharacterDock ? 1 : 3;

    public Thickness InputMargin => HasFrame ? FramedInputMargin : new Thickness(8);

    public int UsersRow => UsesCharacterDock ? 3 : 1;

    public int UsersColumn => UsesCharacterDock ? 0 : 2;

    public int UsersColumnSpan => UsesCharacterDock ? 3 : 1;

    public int UsersRowSpan => HasFrame && !UsesCharacterDock ? 2 : 1;

    public double UsersMaxHeight => UsesCharacterDock ? 190 : double.PositiveInfinity;

    public Thickness UsersMargin => UsesCharacterDock ? new Thickness(8, 2, 8, 8) : new Thickness(0, 4, 8, 8);

    /// <summary>A column splitter only has something to resize while the user list is a column.</summary>
    public bool ShowSplitter => !UsesCharacterDock;

    public int SplitterRowSpan => HasFrame ? 2 : 1;

    public int FrameColumnSpan => UsesCharacterDock ? 3 : 1;

    public Orientation UserListOrientation => UsesCharacterDock ? Orientation.Horizontal : Orientation.Vertical;

    private static Color ToColor(int packed) => Color.FromRgb((byte)(packed >> 16), (byte)(packed >> 8), (byte)packed);

    private static Color Mix(Color from, Color to, double amount) => Color.FromRgb(
        (byte)Math.Round(from.R + (to.R - from.R) * amount),
        (byte)Math.Round(from.G + (to.G - from.G) * amount),
        (byte)Math.Round(from.B + (to.B - from.B) * amount));

    private static LinearGradientBrush Gradient(RelativePoint start, RelativePoint end, Color first, Color middle, Color last) => new()
    {
        StartPoint = start,
        EndPoint = end,
        GradientStops =
        {
            new GradientStop(first, 0),
            new GradientStop(middle, 0.45),
            new GradientStop(last, 1),
        },
    };
}
