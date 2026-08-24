using System.Collections.ObjectModel;
using Avalonia.Controls;
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
                new ProductOption(BncsProduct.Sc2, "StarCraft II", GameIconLoader.Get("sc2")),
                // Dedicated SC:R/WC3:R icons are still pending (the user's own future-icon-work
                // backlog item) — these borrow the closest existing classic icon for the picker
                // specifically. The Friends-list/roster side is a separate, harder limitation:
                // Stimpak's own data has no per-contact game field at all, so every Stimpak-
                // backed contact there falls back to the plain "sc2" icon regardless — see
                // ChatIcon.GetProductIconKey's "sc2" sentinel case.
                new ProductOption(BncsProduct.ScRemastered, "StarCraft: Remastered", GameIconLoader.Get("sc")),
                new ProductOption(BncsProduct.Wc3Reforged, "Warcraft III: Reforged", GameIconLoader.Get("war3")),
            ])
            .ToList();

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(RequiresExpansionKey))]
    [NotifyPropertyChangedFor(nameof(RequiresCdKey))]
    [NotifyPropertyChangedFor(nameof(AllowsOfficialServers))]
    [NotifyPropertyChangedFor(nameof(ServerCompatibilityNote))]
    [NotifyPropertyChangedFor(nameof(ProductIconImage))]
    [NotifyPropertyChangedFor(nameof(IsStimpakBackedProduct))]
    public partial string Product { get; set; }

    public bool IsStimpakBackedProduct => BncsProduct.IsStimpakBacked(Product);

    // --- Tab group icon: which icon (see IconCatalog) this bot's Tab Group shows on its
    // collapsed top-level tab, once it has one — see TabGroupIconStore. Groups aren't a
    // separate persisted entity, so this is keyed by the group name text itself; changing
    // the name switches which assignment is being edited. ---

    private static readonly IconOption NoGroupIconSentinel = new("", "(use first bot's icon)", null);

    public IReadOnlyList<IconOption> AvailableGroupIcons { get; } =
        new[] { NoGroupIconSentinel }
            .Concat(IconCatalog.GameIcons.Concat(IconCatalog.CustomIcons).Concat(IconCatalog.Bnet2Icons)
                .Select(s => new IconOption(s.Key, s.DisplayName, GameIconLoader.Get(s.Key))))
            .ToList();

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsInTabGroup))]
    public partial string TabGroup { get; set; }

    [ObservableProperty]
    public partial IconOption SelectedGroupIcon { get; set; } = NoGroupIconSentinel;

    public bool IsInTabGroup => !string.IsNullOrWhiteSpace(TabGroup);

    // --- StarCraft II login: which saved Battle.net login (see "Manage Battle.net
    // Profiles..." under the Customize menu) this bot uses. The actual sign-in itself
    // happens through Stimpak's own native window on first real connect (or from the
    // Manage Battle.net Profiles window's "Sign In..." action) — picking a profile here
    // just says which cached session to use, sharing one across bots is exactly the point. ---

    /// <summary>Sentinel row appended to AvailableProfiles — picking it reveals the inline "name, then Create" mini-form rather than immediately assigning a profile.</summary>
    private static readonly BattlenetCredentialProfile NewProfileSentinel = new() { Id = "", Name = "+ New Profile..." };

    public ObservableCollection<BattlenetCredentialProfile> AvailableProfiles { get; } = [];

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsCreatingNewProfile))]
    public partial BattlenetCredentialProfile? SelectedProfile { get; set; }

    public bool IsCreatingNewProfile => ReferenceEquals(SelectedProfile, NewProfileSentinel);

    [ObservableProperty]
    public partial string NewProfileName { get; set; } = "";

    partial void OnSelectedProfileChanged(BattlenetCredentialProfile? value)
    {
        if (value is null || ReferenceEquals(value, NewProfileSentinel))
        {
            return;
        }

        Config.BattlenetCredentialProfileId = value.Id;
    }

    [RelayCommand]
    private void CreateProfile()
    {
        if (string.IsNullOrWhiteSpace(NewProfileName))
        {
            return;
        }

        var profile = BattlenetCredentialProfileStore.CreateAndSave(NewProfileName);
        NewProfileName = "";
        RefreshAvailableProfiles(profile.Id);
    }

    private void RefreshAvailableProfiles(string? preferredId = null)
    {
        var selectedId = preferredId ?? SelectedProfile?.Id;
        AvailableProfiles.Clear();
        foreach (var profile in BattlenetCredentialProfileStore.Profiles)
        {
            AvailableProfiles.Add(profile);
        }

        AvailableProfiles.Add(NewProfileSentinel);
        SelectedProfile = AvailableProfiles.FirstOrDefault(p => p.Id == selectedId);
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

    public bool RequiresExpansionKey => !IsStimpakBackedProduct && BncsProduct.RequiresExpansionCdKey(Product);

    public bool RequiresCdKey => !IsStimpakBackedProduct && BncsProduct.RequiresCdKey(Product);

    public bool AllowsOfficialServers => BncsProduct.GetServerCompatibility(Product) == ServerCompatibility.Both;

    public string? ServerCompatibilityNote =>
        BncsProduct.Catalog.TryGetValue(Product, out var info) ? info.Notes : null;

    public Bitmap? ProductIconImage => GameIconLoader.Get(BncsProduct.GetIconKey(Product));

    // --- Official server picker (4 fixed choices + "Other...") ---

    public ObservableCollection<string> OfficialServerOptions { get; } =
        new(BncsProduct.OfficialBattlenetServers.Append(OtherServerOption));

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsOtherServerSelected))]
    public partial string SelectedOfficialServer { get; set; }

    public bool IsOtherServerSelected => SelectedOfficialServer == OtherServerOption;

    /// <summary>
    /// Mirrors Config.BattlenetServer as its own observable VM property (same
    /// reason as ProxyEnabled below) — both custom-server AutoCompleteBoxes
    /// bind here instead of straight to Config.BattlenetServer, so setting a
    /// server from code (see OnBattlenetServerChanged's default-home-channel
    /// lookup) can also push HomeChannel out to its own bound TextBox.
    /// </summary>
    [ObservableProperty]
    public partial string BattlenetServer { get; set; }

    /// <summary>Mirrors Config.HomeChannel — needs to be observable so setting it from code (a per-server default, see OnBattlenetServerChanged) actually updates the bound TextBox, not just the underlying Config.</summary>
    [ObservableProperty]
    public partial string HomeChannel { get; set; }

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

    [ObservableProperty]
    public partial string CustomSchemeName { get; set; }

    /// <summary>Which named schemes from the shared library ("Manage Colors" under the Customize menu) this bot can pick from when ColorScheme is Custom — editing them happens there now, not here.</summary>
    public ObservableCollection<LibrarySchemeOption> LibrarySchemes { get; } = [];

    [ObservableProperty]
    public partial LibrarySchemeOption? SelectedLibraryScheme { get; set; }

    private const string DefaultIconSetLabel = "Default (bundled icons)";

    /// <summary>Which saved IconSetStore set this bot uses — DefaultIconSetLabel (maps to Config.IconSetName = "") plus every set editable/creatable from "Manage Icons..." under the Customize menu.</summary>
    public ObservableCollection<string> AvailableIconSets { get; } = [];

    [ObservableProperty]
    public partial string IconSetName { get; set; }

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
        TabGroup = config.TabGroup;
        SelectedGroupIcon = AvailableGroupIcons.FirstOrDefault(o => o.Key == TabGroupIconStore.GetIconKey(TabGroup)) ?? NoGroupIconSentinel;
        ProxyEnabled = config.ProxyEnabled;
        ProxyProtocol = config.ProxyProtocol;
        SelectedOfficialServer = BncsProduct.OfficialBattlenetServers.Contains(config.BattlenetServer)
            ? config.BattlenetServer
            : OtherServerOption;
        BattlenetServer = config.BattlenetServer;
        HomeChannel = config.HomeChannel;
        ColorScheme = config.ChatColorScheme;
        CustomSchemeName = config.CustomColorSchemeName;
        IconSetName = string.IsNullOrEmpty(config.IconSetName) ? DefaultIconSetLabel : config.IconSetName;
        RefreshServerSuggestions();
        RefreshLibrarySchemes();
        RefreshAvailableIconSets();
        RefreshAvailableProfiles(config.BattlenetCredentialProfileId);
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

    private void RefreshAvailableIconSets()
    {
        var selected = IconSetName;
        AvailableIconSets.Clear();
        AvailableIconSets.Add(DefaultIconSetLabel);
        foreach (var name in IconSetStore.ListSets())
        {
            AvailableIconSets.Add(name);
        }

        IconSetName = AvailableIconSets.Contains(selected) ? selected : DefaultIconSetLabel;
    }

    partial void OnIconSetNameChanged(string value) =>
        Config.IconSetName = value == DefaultIconSetLabel ? "" : value;

    partial void OnProductChanged(string value)
    {
        Config.Product = value;
        RefreshServerSuggestions();
    }

    partial void OnProxyProtocolChanged(ProxyProtocol value) => Config.ProxyProtocol = value;

    partial void OnTabGroupChanged(string value)
    {
        Config.TabGroup = value;
        SelectedGroupIcon = AvailableGroupIcons.FirstOrDefault(o => o.Key == TabGroupIconStore.GetIconKey(value)) ?? NoGroupIconSentinel;
    }

    /// <summary>Applied immediately (not batched into Save) since a group's icon is shared state keyed by name, not part of this one bot's own Config.</summary>
    partial void OnSelectedGroupIconChanged(IconOption value)
    {
        if (!string.IsNullOrWhiteSpace(TabGroup))
        {
            TabGroupIconStore.SetIconKey(TabGroup, value.Key);
        }
    }

    partial void OnProxyEnabledChanged(bool value) => Config.ProxyEnabled = value;

    partial void OnSelectedOfficialServerChanged(string value)
    {
        if (value != OtherServerOption)
        {
            BattlenetServer = value;
        }
    }

    /// <summary>Applies a known-good default HomeChannel for specific public servers (currently just atlas.bnetdocs.org → "Town Square") the first time that server's picked — never overwrites a channel the user already typed.</summary>
    partial void OnBattlenetServerChanged(string value)
    {
        Config.BattlenetServer = value;
        if (string.IsNullOrWhiteSpace(HomeChannel) && BncsProduct.GetDefaultHomeChannel(value) is { } suggested)
        {
            HomeChannel = suggested;
        }
    }

    partial void OnHomeChannelChanged(string value) => Config.HomeChannel = value;

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

/// <summary>One pickable entry in a "choose an icon" list — see ConfigViewModel.AvailableGroupIcons.</summary>
public sealed record IconOption(string Key, string DisplayName, Bitmap? Icon);

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
