using System.Collections.ObjectModel;
using Avalonia.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;

namespace Invigoration.App.ViewModels;

/// <summary>One labeled option in a theme-editor dropdown.</summary>
public sealed record ThemeChoice<T>(T Value, string Label);

/// <summary>One of a theme's seven material colors, as an editable color.</summary>
public sealed partial class ThemeMaterialSlotViewModel(string label, string hint, Color value, Action onChanged) : ObservableObject
{
    public string Label { get; } = label;

    public string Hint { get; } = hint;

    [ObservableProperty]
    public partial Color Value { get; set; } = value;

    partial void OnValueChanged(Color value) => onChanged();

    public int Packed => (Value.R << 16) | (Value.G << 8) | Value.B;
}

/// <summary>A line of sample chat for the preview, already colored from the theme's palette.</summary>
public sealed record ThemePreviewLine(string Name, IBrush NameBrush, string Text, IBrush TextBrush);

/// <summary>
/// Customize → Manage Themes: every theme (ThemeLibrary), an editor for the selected one, and a
/// live preview drawn with the same ThemeFrame/ThemeDivider/ThemeStyles a bot's tab uses, so what
/// you see here is what a bot gets. Built-in themes are read-only — Duplicate makes an editable
/// copy. Saving a custom theme redraws every bot already using it (ThemeLibrary.ThemesChanged).
/// </summary>
public sealed partial class ThemeManagerViewModel : ObservableObject, IThemedSurface
{
    private bool _loading;

    public ThemeManagerViewModel()
    {
        MaterialSlots = [];
        ReloadThemes(ThemeLibrary.DiabloIIId);
    }

    public ObservableCollection<ChatTheme> Themes { get; } = [];

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsEditable))]
    [NotifyPropertyChangedFor(nameof(ReadOnlyNote))]
    [NotifyCanExecuteChangedFor(nameof(SaveCommand))]
    [NotifyCanExecuteChangedFor(nameof(DeleteCommand))]
    [NotifyCanExecuteChangedFor(nameof(DuplicateCommand))]
    public partial ChatTheme? SelectedTheme { get; set; }

    public bool IsEditable => SelectedTheme is { IsBuiltIn: false };

    public string ReadOnlyNote => SelectedTheme is { IsBuiltIn: true }
        ? "Built-in themes can't be changed. Duplicate this one to make an editable copy."
        : "";

    public IReadOnlyList<ThemeChoice<ThemeFrameStyle>> FrameChoices { get; } =
    [
        new(ThemeFrameStyle.None, "None (plain)"),
        new(ThemeFrameStyle.DiabloII, "Diablo II: carved stone and gold"),
        new(ThemeFrameStyle.StarCraft, "StarCraft: steel console"),
        new(ThemeFrameStyle.Warcraft, "Warcraft: oak and iron"),
    ];

    public IReadOnlyList<ThemeChoice<ThemeLayout>> LayoutChoices { get; } =
    [
        new(ThemeLayout.Standard, "User list on the right"),
        new(ThemeLayout.CharacterDock, "Character portraits along the bottom"),
    ];

    public IReadOnlyList<ThemeChoice<ChatColorScheme?>> ColorChoices { get; } =
    [
        new(null, "Leave each bot's colors alone"),
        new(ChatColorScheme.Invigoration, "Invigoration (classic)"),
        new(ChatColorScheme.StarCraft, "BNU`Bot StarCraft"),
        new(ChatColorScheme.DiabloII, "BNU`Bot Diablo"),
        new(ChatColorScheme.Warcraft, "Warcraft"),
    ];

    [ObservableProperty]
    public partial string Name { get; set; } = "";

    [ObservableProperty]
    public partial ThemeChoice<ThemeFrameStyle>? Frame { get; set; }

    [ObservableProperty]
    public partial ThemeChoice<ThemeLayout>? Layout { get; set; }

    [ObservableProperty]
    public partial ThemeChoice<ChatColorScheme?>? Colors { get; set; }

    [ObservableProperty]
    public partial bool ShowChatGem { get; set; }

    [ObservableProperty]
    public partial bool ShowNamePlate { get; set; }

    [ObservableProperty]
    public partial string HeaderFont { get; set; } = "";

    public ObservableCollection<ThemeMaterialSlotViewModel> MaterialSlots { get; }

    [ObservableProperty]
    public partial string StatusText { get; set; } = "";

    /// <summary>The preview's theme, rebuilt from the editor on every change.</summary>
    [ObservableProperty]
    public partial ChatThemeViewModel Theme { get; set; } = new(ThemeLibrary.BuiltIns[0]);

    [ObservableProperty]
    public partial IBrush PreviewBackground { get; set; } = Brushes.Black;

    [ObservableProperty]
    public partial IReadOnlyList<ThemePreviewLine> PreviewLines { get; set; } = [];

    [ObservableProperty]
    public partial IBrush PreviewGemBrush { get; set; } = Brushes.Blue;

    partial void OnSelectedThemeChanged(ChatTheme? value)
    {
        if (value is not null)
        {
            LoadEditor(value);
        }
    }

    partial void OnNameChanged(string value) => EditorChanged();

    partial void OnFrameChanged(ThemeChoice<ThemeFrameStyle>? value) => EditorChanged();

    partial void OnLayoutChanged(ThemeChoice<ThemeLayout>? value) => EditorChanged();

    partial void OnColorsChanged(ThemeChoice<ChatColorScheme?>? value) => EditorChanged();

    partial void OnShowChatGemChanged(bool value) => EditorChanged();

    partial void OnShowNamePlateChanged(bool value) => EditorChanged();

    partial void OnHeaderFontChanged(string value) => EditorChanged();

    /// <summary>The editor's current contents as a theme, keeping the selected theme's id.</summary>
    public ChatTheme BuildTheme() => new()
    {
        Id = SelectedTheme?.Id ?? "",
        Name = string.IsNullOrWhiteSpace(Name) ? "Untitled Theme" : Name.Trim(),
        Frame = Frame?.Value ?? ThemeFrameStyle.None,
        Layout = Layout?.Value ?? ThemeLayout.Standard,
        ColorScheme = Colors?.Value,
        ShowChatGem = ShowChatGem,
        ShowNamePlate = ShowNamePlate,
        HeaderFont = HeaderFont.Trim(),
        Materials = new ThemeMaterials
        {
            Band = MaterialSlots[0].Packed,
            BandShade = MaterialSlots[1].Packed,
            Trim = MaterialSlots[2].Packed,
            TrimShade = MaterialSlots[3].Packed,
            Metal = MaterialSlots[4].Packed,
            Accent = MaterialSlots[5].Packed,
            Well = MaterialSlots[6].Packed,
        },
    };

    [RelayCommand(CanExecute = nameof(CanDuplicate))]
    private void Duplicate()
    {
        var source = SelectedTheme!;
        var copy = ThemeLibrary.Duplicate(source.IsBuiltIn ? source : BuildTheme(), $"{source.Name} (copy)");
        ThemeLibrary.Save(copy);
        ReloadThemes(copy.Id);
        StatusText = $"Made \"{copy.Name}\". Edit it, then Save.";
    }

    private bool CanDuplicate() => SelectedTheme is not null;

    [RelayCommand(CanExecute = nameof(IsEditable))]
    private void Save()
    {
        var theme = BuildTheme();
        ThemeLibrary.Save(theme);
        ReloadThemes(theme.Id);
        StatusText = $"Saved \"{theme.Name}\". Bots using it have been updated.";
    }

    [RelayCommand(CanExecute = nameof(IsEditable))]
    private void Delete()
    {
        var theme = SelectedTheme!;
        ThemeLibrary.Delete(theme.Id);
        ReloadThemes(ThemeLibrary.DefaultId);
        StatusText = $"Deleted \"{theme.Name}\". Any bot that used it is back on Default.";
    }

    private void ReloadThemes(string selectId)
    {
        Themes.Clear();
        foreach (var theme in ThemeLibrary.All())
        {
            Themes.Add(theme);
        }

        SelectedTheme = Themes.FirstOrDefault(t => t.Id == selectId) ?? Themes[0];
    }

    private void LoadEditor(ChatTheme theme)
    {
        _loading = true;
        try
        {
            Name = theme.Name;
            Frame = FrameChoices.First(c => c.Value == theme.Frame);
            Layout = LayoutChoices.First(c => c.Value == theme.Layout);
            Colors = ColorChoices.First(c => c.Value == theme.ColorScheme);
            ShowChatGem = theme.ShowChatGem;
            ShowNamePlate = theme.ShowNamePlate;
            HeaderFont = theme.HeaderFont;

            var m = theme.Materials;
            MaterialSlots.Clear();
            MaterialSlots.Add(new("Frame body", "Stone, steel or wood, lit side", ToColor(m.Band), EditorChanged));
            MaterialSlots.Add(new("Frame body shadow", "The same, shadowed side", ToColor(m.BandShade), EditorChanged));
            MaterialSlots.Add(new("Trim", "Inlay and edges, lit", ToColor(m.Trim), EditorChanged));
            MaterialSlots.Add(new("Trim shadow", "Inlay and edges, shadowed", ToColor(m.TrimShade), EditorChanged));
            MaterialSlots.Add(new("Hardware", "Rivets, bolts and brackets", ToColor(m.Metal), EditorChanged));
            MaterialSlots.Add(new("Accent", "Channel name, Send label, glow", ToColor(m.Accent), EditorChanged));
            MaterialSlots.Add(new("Text box", "Behind what you type", ToColor(m.Well), EditorChanged));
        }
        finally
        {
            _loading = false;
        }

        StatusText = "";
        RefreshPreview();
    }

    private void EditorChanged()
    {
        if (_loading)
        {
            return;
        }

        if (IsEditable)
        {
            StatusText = "Unsaved changes.";
        }

        RefreshPreview();
    }

    private void RefreshPreview()
    {
        if (MaterialSlots.Count < 7)
        {
            return;
        }

        var theme = BuildTheme();
        Theme = new ChatThemeViewModel(theme) { ChannelName = "Town Square" };

        var palette = (theme.ColorScheme ?? ChatColorScheme.Invigoration) switch
        {
            ChatColorScheme.StarCraft => ChatPalette.StarCraft,
            ChatColorScheme.DiabloII => ChatPalette.DiabloII,
            ChatColorScheme.Warcraft => ChatPalette.Warcraft,
            _ => ChatPalette.Invigoration,
        };

        PreviewBackground = Brush(palette.Background);
        PreviewGemBrush = Brush(palette.Blue);
        PreviewLines =
        [
            new("", Brushes.Transparent, "Joining channel: Town Square", Brush(palette.Channel)),
            new("Zeal: ", Brush(palette.UserNameDefault), "anyone up for a Baal run?", Brush(palette.White)),
            new("Tagban: ", Brush(palette.SelfUserName), "give me five minutes, grabbing my gear", Brush(palette.White)),
            new("", Brushes.Transparent, "Bonecrusher has joined the channel.", Brush(palette.Gray)),
            new("", Brushes.Transparent, "You whisper to Zeal: bring the runes", Brush(palette.Whisper)),
        ];
    }

    private static Color ToColor(int packed) => Color.FromRgb((byte)(packed >> 16), (byte)(packed >> 8), (byte)packed);

    private static IBrush Brush(RgbColor c) => new SolidColorBrush(Color.FromRgb(c.R, c.G, c.B));
}
