using System.Collections.ObjectModel;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.Models;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;
using Invigoration.Core.Networking;
using Invigoration.Core.Protocol;

namespace Invigoration.App.ViewModels;

/// <summary>
/// Wraps a BotConfig being edited so the config window can react to product
/// selection: which server field(s) to show, whether a second (expansion)
/// CD-key field is needed, and which server suggestions are offered.
/// </summary>
public partial class ConfigViewModel : ObservableObject
{
    private const string OtherServerOption = "Other...";

    public BotConfig Config { get; }

    public IReadOnlyList<ProductOption> AvailableProducts { get; } =
        BncsProduct.Catalog.Values
            .Select(info => new ProductOption(info.WireCode, info.DisplayName, GameIconLoader.Get(info.IconKey)))
            .OrderBy(p => p.DisplayName)
            .Concat(
            [
                new ProductOption("SC2", "StarCraft II (coming soon)", null, IsSelectable: false),
                new ProductOption("SCRM", "StarCraft: Remastered (coming soon)", null, IsSelectable: false),
            ])
            .ToList();

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(RequiresExpansionKey))]
    [NotifyPropertyChangedFor(nameof(RequiresCdKey))]
    [NotifyPropertyChangedFor(nameof(AllowsOfficialServers))]
    [NotifyPropertyChangedFor(nameof(ServerCompatibilityNote))]
    [NotifyPropertyChangedFor(nameof(ProductIconImage))]
    public partial string Product { get; set; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(AllowsOfficialServers))]
    [NotifyPropertyChangedFor(nameof(IsBncsBinary))]
    [NotifyPropertyChangedFor(nameof(IsTelnetGateway))]
    public partial ConnectionMode ConnectionMode { get; set; }

    public bool IsBncsBinary
    {
        get => ConnectionMode == ConnectionMode.BncsBinary;
        set
        {
            if (value)
            {
                ConnectionMode = ConnectionMode.BncsBinary;
            }
        }
    }

    public bool IsTelnetGateway
    {
        get => ConnectionMode == ConnectionMode.TelnetGateway;
        set
        {
            if (value)
            {
                ConnectionMode = ConnectionMode.TelnetGateway;
            }
        }
    }

    // --- Proxy ---

    /// <summary>
    /// Mirrors Config.ProxyEnabled as its own observable VM property (rather
    /// than binding views directly to Config.ProxyEnabled) because BotConfig
    /// is a plain class with no INotifyPropertyChanged — a view bound
    /// straight to Config.ProxyEnabled would write fine but never be told to
    /// re-read it, so e.g. an IsVisible toggle for the rest of the proxy
    /// fields would never actually show/hide when the checkbox changes.
    /// </summary>
    [ObservableProperty]
    public partial bool ProxyEnabled { get; set; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsSocks5Proxy))]
    [NotifyPropertyChangedFor(nameof(IsHttpProxy))]
    public partial ProxyProtocol ProxyProtocol { get; set; }

    public bool IsSocks5Proxy
    {
        get => ProxyProtocol == ProxyProtocol.Socks5;
        set
        {
            if (value)
            {
                ProxyProtocol = ProxyProtocol.Socks5;
            }
        }
    }

    public bool IsHttpProxy
    {
        get => ProxyProtocol == ProxyProtocol.Http;
        set
        {
            if (value)
            {
                ProxyProtocol = ProxyProtocol.Http;
            }
        }
    }

    public ObservableCollection<string> ServerSuggestions { get; } = [];

    public bool RequiresExpansionKey => BncsProduct.RequiresExpansionCdKey(Product);

    public bool RequiresCdKey => BncsProduct.RequiresCdKey(Product);

    public bool AllowsOfficialServers =>
        ConnectionMode == ConnectionMode.BncsBinary &&
        BncsProduct.GetServerCompatibility(Product) == ServerCompatibility.Both;

    public string? ServerCompatibilityNote =>
        BncsProduct.Catalog.TryGetValue(Product, out var info) ? info.Notes : null;

    public Bitmap? ProductIconImage => GameIconLoader.Get(
        BncsProduct.Catalog.TryGetValue(Product, out var info) ? info.IconKey : ChatIcon.GetProductIconKey(Product));

    // --- Official server picker (4 fixed choices + "Other...") ---

    public ObservableCollection<string> OfficialServerOptions { get; } =
        new(BncsProduct.OfficialBattlenetServers.Append(OtherServerOption));

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsOtherServerSelected))]
    public partial string SelectedOfficialServer { get; set; }

    public bool IsOtherServerSelected => SelectedOfficialServer == OtherServerOption;

    // --- Chat color scheme ---

    public IReadOnlyList<ColorSchemeOption> AvailableColorSchemes { get; } =
    [
        new(ChatColorScheme.Invigoration, "Invigoration (classic)"),
        new(ChatColorScheme.StarCraft, "BNU`Bot StarCraft"),
        new(ChatColorScheme.DiabloII, "BNU`Bot Diablo"),
        new(ChatColorScheme.Custom, "Custom..."),
    ];

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(PaletteSwatches))]
    [NotifyPropertyChangedFor(nameof(IsCustomScheme))]
    public partial ChatColorScheme ColorScheme { get; set; }

    public bool IsCustomScheme => ColorScheme == ChatColorScheme.Custom;

    public ObservableCollection<CustomColorSlotViewModel> CustomColorSlots { get; }

    [ObservableProperty]
    public partial string CustomSchemeName { get; set; }

    public ObservableCollection<LibrarySchemeOption> LibrarySchemes { get; } = [];

    [ObservableProperty]
    public partial LibrarySchemeOption? SelectedLibraryScheme { get; set; }

    public IReadOnlyList<PaletteSwatch> PaletteSwatches
    {
        get
        {
            var p = ChatPalette.ForScheme(Config);
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

    public ConfigViewModel(BotConfig config)
    {
        Config = config;
        Product = config.Product;
        ConnectionMode = config.ConnectionMode;
        ProxyEnabled = config.ProxyEnabled;
        ProxyProtocol = config.ProxyProtocol;
        SelectedOfficialServer = BncsProduct.OfficialBattlenetServers.Contains(config.BattlenetServer)
            ? config.BattlenetServer
            : OtherServerOption;
        ColorScheme = config.ChatColorScheme;
        var c = config.CustomColors;
        CustomColorSlots = new ObservableCollection<CustomColorSlotViewModel>
        {
            new("Background", () => c.Background, v => c.Background = v, RefreshPaletteSwatches),
            new("Talk (default text)", () => c.White, v => c.White = v, RefreshPaletteSwatches),
            new("Channel joined", () => c.Channel, v => c.Channel = v, RefreshPaletteSwatches),
            new("Info / status", () => c.Info, v => c.Info = v, RefreshPaletteSwatches),
            new("Error / warning", () => c.Error, v => c.Error = v, RefreshPaletteSwatches),
            new("Debug", () => c.Debug, v => c.Debug = v, RefreshPaletteSwatches),
            new("Join / leave", () => c.Gray, v => c.Gray = v, RefreshPaletteSwatches),
            new("Your own name", () => c.SelfUserName, v => c.SelfUserName = v, RefreshPaletteSwatches),
            new("Whisper / ignored", () => c.Whisper, v => c.Whisper = v, RefreshPaletteSwatches),
            new("Highlight (active tab)", () => c.Highlight, v => c.Highlight = v, RefreshPaletteSwatches),
            new("Ignored user's name", () => c.Red, v => c.Red = v, RefreshPaletteSwatches),
            new("Bnet-rep name/chat", () => c.Green, v => c.Green = v, RefreshPaletteSwatches),
            new("Blizzard-rep name/chat", () => c.Cyan, v => c.Cyan = v, RefreshPaletteSwatches),
            new("Speaker name/chat", () => c.Speaker, v => c.Speaker = v, RefreshPaletteSwatches),
            new("Guest name/chat", () => c.Guest, v => c.Guest = v, RefreshPaletteSwatches),
            new("Default username", () => c.UserNameDefault, v => c.UserNameDefault = v, RefreshPaletteSwatches),
            new("Default emote", () => c.EmoteDefault, v => c.EmoteDefault = v, RefreshPaletteSwatches),
        };
        CustomSchemeName = config.CustomColorSchemeName;
        RefreshServerSuggestions();
        RefreshLibrarySchemes();
    }

    private void RefreshPaletteSwatches() => OnPropertyChanged(nameof(PaletteSwatches));

    partial void OnCustomSchemeNameChanged(string value) => Config.CustomColorSchemeName = value;

    private void RefreshLibrarySchemes()
    {
        LibrarySchemes.Clear();
        foreach (var (filePath, name) in ColorSchemeLibrary.ListSchemes())
        {
            LibrarySchemes.Add(new LibrarySchemeOption(filePath, name));
        }
    }

    /// <summary>Loads a scheme (from the library dropdown or an imported file) into the live editor: switches to Custom, overwrites every color, and refreshes the pickers/preview.</summary>
    private void ApplyNamedPalette(NamedCustomPalette scheme)
    {
        CustomSchemeName = scheme.Name;
        var dst = Config.CustomColors;
        var src = scheme.Colors;
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

        ColorScheme = ChatColorScheme.Custom;
        foreach (var slot in CustomColorSlots)
        {
            slot.Refresh();
        }

        RefreshPaletteSwatches();
    }

    [RelayCommand]
    private void LoadLibraryScheme()
    {
        if (SelectedLibraryScheme is not { } selected)
        {
            return;
        }

        ApplyNamedPalette(ColorSchemeLibrary.Load(selected.FilePath));
    }

    [RelayCommand]
    private void SaveSchemeToLibrary()
    {
        ColorSchemeLibrary.Save(new NamedCustomPalette { Name = CustomSchemeName, Colors = Config.CustomColors });
        RefreshLibrarySchemes();
    }

    /// <summary>Loads a scheme file picked from outside the library (e.g. a friend's email attachment), and adds it to the library so it's available going forward too.</summary>
    public void ImportSchemeFile(string filePath)
    {
        var imported = ColorSchemeLibrary.Load(filePath);
        ColorSchemeLibrary.Save(imported);
        RefreshLibrarySchemes();
        ApplyNamedPalette(imported);
    }

    partial void OnProductChanged(string value)
    {
        Config.Product = value;
        RefreshServerSuggestions();
    }

    partial void OnConnectionModeChanged(ConnectionMode value)
    {
        Config.ConnectionMode = value;
        RefreshServerSuggestions();
    }

    partial void OnProxyProtocolChanged(ProxyProtocol value) => Config.ProxyProtocol = value;

    partial void OnProxyEnabledChanged(bool value) => Config.ProxyEnabled = value;

    partial void OnSelectedOfficialServerChanged(string value)
    {
        if (value != OtherServerOption)
        {
            Config.BattlenetServer = value;
        }
    }

    partial void OnColorSchemeChanged(ChatColorScheme value) => Config.ChatColorScheme = value;

    private void RefreshServerSuggestions()
    {
        ServerSuggestions.Clear();
        foreach (var server in BncsProduct.SuggestedPrivateServers)
        {
            ServerSuggestions.Add(server);
        }
    }
}

public sealed record ProductOption(string WireCode, string DisplayName, Bitmap? Icon, bool IsSelectable = true)
{
    public double DisplayOpacity => IsSelectable ? 1.0 : 0.4;
}

public sealed record ColorSchemeOption(ChatColorScheme Value, string DisplayName);

/// <summary>One scheme file found in the Colors library folder.</summary>
public sealed record LibrarySchemeOption(string FilePath, string Name);

public sealed record PaletteSwatch(string Label, RgbColor Color)
{
    public IBrush Brush { get; } = new SolidColorBrush(Avalonia.Media.Color.FromRgb(Color.R, Color.G, Color.B));
}

/// <summary>
/// One editable role in a Custom color scheme: reads/writes a single 0xRRGGBB
/// packed int on BotConfig.CustomColors via the getter/setter it's given,
/// exposed as an Avalonia Color for a ColorPicker to bind to.
/// </summary>
public partial class CustomColorSlotViewModel : ObservableObject
{
    private readonly Func<int> _getPacked;
    private readonly Action<int> _setPacked;
    private readonly Action _onChanged;
    private bool _suppressWriteback;

    public string Label { get; }

    [ObservableProperty]
    public partial Color Value { get; set; }

    public CustomColorSlotViewModel(string label, Func<int> getPacked, Action<int> setPacked, Action onChanged)
    {
        Label = label;
        // Assigned before Value so OnValueChanged's no-op write-back below has somewhere to go.
        _getPacked = getPacked;
        _setPacked = setPacked;
        _onChanged = onChanged;
        Refresh();
    }

    /// <summary>Re-reads the underlying value (e.g. after an imported scheme overwrote it directly) without re-writing it back.</summary>
    public void Refresh()
    {
        var packed = _getPacked();
        _suppressWriteback = true;
        Value = Color.FromRgb((byte)(packed >> 16), (byte)(packed >> 8), (byte)packed);
        _suppressWriteback = false;
    }

    partial void OnValueChanged(Color value)
    {
        if (_suppressWriteback)
        {
            return;
        }

        _setPacked((value.R << 16) | (value.G << 8) | value.B);
        _onChanged();
    }
}
