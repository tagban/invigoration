using System.Collections.ObjectModel;
using Avalonia.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;

namespace Invigoration.App.ViewModels;

/// <summary>
/// Shared (cross-bot) color-scheme editor, reachable from the top-level
/// Customize menu — full swatch-by-swatch editing lives here now, extracted
/// out of the per-bot Config window's Appearance section, which is just a
/// picker ("which saved scheme does this bot use") now. Operates on its own
/// working CustomChatPalette rather than any specific BotConfig, since a
/// saved scheme in the shared ColorSchemeLibrary isn't owned by one bot.
/// </summary>
public partial class ColorManagerViewModel : ObservableObject
{
    private CustomChatPalette _working = new();

    public ObservableCollection<LibrarySchemeOption> Schemes { get; } = [];

    [ObservableProperty]
    public partial LibrarySchemeOption? SelectedScheme { get; set; }

    [ObservableProperty]
    public partial string SchemeName { get; set; } = "New Scheme";

    public ObservableCollection<CustomColorSlotViewModel> ColorSlots { get; }

    public IReadOnlyList<PaletteSwatch> PaletteSwatches
    {
        get
        {
            var p = ChatPalette.FromCustom(_working);
            return
            [
                new("Background", p.Background),
                new("Talk", p.White),
                new("Channel", p.Channel),
                new("Info", p.Info),
                new("Error", p.Error),
                new("Debug", p.Debug),
                new("Whisper", p.Whisper),
                new("Emote", p.GetEmoteColor(0)),
                new("Highlight", p.Highlight),
                new("Operator", p.GetUserNameColor((uint)UserFlags.Operator)),
                new("Speaker", p.GetUserNameColor((uint)UserFlags.Speaker)),
                new("Rep", p.GetUserNameColor((uint)UserFlags.Blizzard)),
                new("Guest", p.GetUserNameColor((uint)UserFlags.Special)),
                new("Ignored", p.GetUserNameColor((uint)UserFlags.Squelched)),
            ];
        }
    }

    public ColorManagerViewModel()
    {
        var c = _working;
        ColorSlots =
        [
            new("Background", () => c.Background, v => c.Background = v, RefreshPreview),
            new("Talk (default text)", () => c.White, v => c.White = v, RefreshPreview),
            new("Channel joined", () => c.Channel, v => c.Channel = v, RefreshPreview),
            new("Info / status", () => c.Info, v => c.Info = v, RefreshPreview),
            new("Error / warning", () => c.Error, v => c.Error = v, RefreshPreview),
            new("Debug", () => c.Debug, v => c.Debug = v, RefreshPreview),
            new("Join / leave", () => c.Gray, v => c.Gray = v, RefreshPreview),
            new("Your own name", () => c.SelfUserName, v => c.SelfUserName = v, RefreshPreview),
            new("Whisper / ignored", () => c.Whisper, v => c.Whisper = v, RefreshPreview),
            new("Highlight (active tab)", () => c.Highlight, v => c.Highlight = v, RefreshPreview),
            new("Ignored user's name", () => c.Red, v => c.Red = v, RefreshPreview),
            new("Bnet-rep name/chat", () => c.Green, v => c.Green = v, RefreshPreview),
            new("Blizzard-rep name/chat", () => c.Cyan, v => c.Cyan = v, RefreshPreview),
            new("Speaker name/chat", () => c.Speaker, v => c.Speaker = v, RefreshPreview),
            new("Guest name/chat", () => c.Guest, v => c.Guest = v, RefreshPreview),
            new("Default username", () => c.UserNameDefault, v => c.UserNameDefault = v, RefreshPreview),
            new("Default emote", () => c.EmoteDefault, v => c.EmoteDefault = v, RefreshPreview),
        ];

        RefreshSchemes();
        if (Schemes.Count > 0)
        {
            SelectedScheme = Schemes[0];
        }
    }

    private void RefreshPreview() => OnPropertyChanged(nameof(PaletteSwatches));

    private void RefreshSchemes()
    {
        var previouslySelected = SelectedScheme;
        Schemes.Clear();
        foreach (var (filePath, name) in ColorSchemeLibrary.ListSchemes())
        {
            Schemes.Add(new LibrarySchemeOption(filePath, name));
        }

        SelectedScheme = previouslySelected is not null
            ? Schemes.FirstOrDefault(s => s.FilePath == previouslySelected.FilePath)
            : null;
    }

    partial void OnSelectedSchemeChanged(LibrarySchemeOption? value)
    {
        if (value is null)
        {
            return;
        }

        var loaded = ColorSchemeLibrary.Load(value.FilePath);
        SchemeName = loaded.Name;
        CopyInto(_working, loaded.Colors);
        foreach (var slot in ColorSlots)
        {
            slot.Refresh();
        }

        RefreshPreview();
    }

    [RelayCommand]
    private void NewScheme()
    {
        SelectedScheme = null;
        SchemeName = "New Scheme";
        CopyInto(_working, new CustomChatPalette());
        foreach (var slot in ColorSlots)
        {
            slot.Refresh();
        }

        RefreshPreview();
    }

    [RelayCommand]
    private void Save()
    {
        if (string.IsNullOrWhiteSpace(SchemeName))
        {
            return;
        }

        var path = ColorSchemeLibrary.Save(new NamedCustomPalette { Name = SchemeName, Colors = _working });
        RefreshSchemes();
        SelectedScheme = Schemes.FirstOrDefault(s => s.FilePath == path);
    }

    [RelayCommand]
    private void Delete()
    {
        if (SelectedScheme is not { } selected)
        {
            return;
        }

        ColorSchemeLibrary.Delete(selected.FilePath);
        RefreshSchemes();
    }

    /// <summary>Loads a scheme file picked from outside the library (e.g. a friend's email attachment), and adds it to the library so it's available going forward too.</summary>
    public void ImportSchemeFile(string filePath)
    {
        var imported = ColorSchemeLibrary.Load(filePath);
        var path = ColorSchemeLibrary.Save(imported);
        RefreshSchemes();
        SelectedScheme = Schemes.FirstOrDefault(s => s.FilePath == path);
    }

    private static void CopyInto(CustomChatPalette dst, CustomChatPalette src)
    {
        dst.Background = src.Background;
        dst.White = src.White;
        dst.Channel = src.Channel;
        dst.Info = src.Info;
        dst.Error = src.Error;
        dst.Debug = src.Debug;
        dst.Gray = src.Gray;
        dst.SelfUserName = src.SelfUserName;
        dst.Whisper = src.Whisper;
        dst.Highlight = src.Highlight;
        dst.Red = src.Red;
        dst.Green = src.Green;
        dst.Cyan = src.Cyan;
        dst.Speaker = src.Speaker;
        dst.Guest = src.Guest;
        dst.UserNameDefault = src.UserNameDefault;
        dst.EmoteDefault = src.EmoteDefault;
    }
}
