using System.Text.Json.Serialization;
using Invigoration.Core.Chat;

namespace Invigoration.Core.Config;

/// <summary>Which ornamental frame a theme draws around a bot's chat and input.</summary>
public enum ThemeFrameStyle
{
    /// <summary>No frame — the plain Invigoration look.</summary>
    None,

    /// <summary>Carved stone with a tarnished-gold inlay, rivet blocks and a channel-name plate, after Diablo II's chat screen.</summary>
    DiabloII,

    /// <summary>Beveled steel console with chamfered corners, status lights and a glowing accent line, after StarCraft's Terran interface.</summary>
    StarCraft,

    /// <summary>Oak planks bound by iron brackets and gold trim, with a hanging name sign, after Warcraft's menus.</summary>
    Warcraft,
}

/// <summary>Where a themed bot's Users/Friends/Clan panel sits.</summary>
public enum ThemeLayout
{
    /// <summary>Down the right-hand side, as a vertical list.</summary>
    Standard,

    /// <summary>Along the bottom as a strip of character portraits — Diablo II's lobby arrangement.</summary>
    CharacterDock,
}

/// <summary>
/// The colors a theme's frame is built from, each a packed 0xRRGGBB like CustomChatPalette. Every
/// frame design draws from the same seven roles, so any design can be recolored in the theme
/// manager without knowing its shapes.
/// </summary>
public sealed class ThemeMaterials
{
    /// <summary>The frame body at its lit end (stone, steel, wood).</summary>
    public int Band { get; set; }

    /// <summary>The frame body at its shadowed end.</summary>
    public int BandShade { get; set; }

    /// <summary>Inlay and edge trim at its lit end.</summary>
    public int Trim { get; set; }

    /// <summary>Inlay and edge trim at its shadowed end.</summary>
    public int TrimShade { get; set; }

    /// <summary>Hardware: rivets, bolts, studs and brackets.</summary>
    public int Metal { get; set; }

    /// <summary>The channel name, Send label, and anything that glows.</summary>
    public int Accent { get; set; }

    /// <summary>The text box's background.</summary>
    public int Well { get; set; }

    public ThemeMaterials Clone() => (ThemeMaterials)MemberwiseClone();
}

/// <summary>
/// A named look for a bot's tab: frame design, where the user list sits, the chat color scheme,
/// and the frame's colors. Each bot picks one (BotConfig.ThemeId); the built-in themes ship with
/// the app and custom ones live as files in the theme library — see <see cref="ThemeLibrary"/>.
/// </summary>
public sealed class ChatTheme
{
    public string Id { get; set; } = "";

    public string Name { get; set; } = "";

    /// <summary>Shipped with the app rather than loaded from the library, so it can't be edited or deleted — only duplicated.</summary>
    [JsonIgnore]
    public bool IsBuiltIn { get; init; }

    public ThemeFrameStyle Frame { get; set; }

    public ThemeLayout Layout { get; set; }

    /// <summary>The chat colors a bot switches to when this theme is picked for it, or null to leave that bot's colors alone. Only applied at the moment of picking — the bot's color scheme stays its own setting afterwards.</summary>
    public ChatColorScheme? ColorScheme { get; set; }

    /// <summary>Diablo II's clickable chat gem, set beside the Send button.</summary>
    public bool ShowChatGem { get; set; }

    /// <summary>The channel name on a plate across the top of the frame.</summary>
    public bool ShowNamePlate { get; set; } = true;

    /// <summary>Font for the name plate and Send button; a comma-separated fallback list. Blank uses the app's normal font.</summary>
    public string HeaderFont { get; set; } = "";

    public ThemeMaterials Materials { get; set; } = new();

    public ChatTheme Clone()
    {
        var copy = (ChatTheme)MemberwiseClone();
        copy.Materials = Materials.Clone();
        return copy;
    }
}
