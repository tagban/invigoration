using System.Text.Json;
using System.Text.Json.Serialization;
using Invigoration.Core.Chat;

namespace Invigoration.Core.Config;

/// <summary>
/// Every theme a bot can pick: the four built-ins (Default, Diablo II, StarCraft, Warcraft), which
/// live in code so they update with the app, plus custom themes saved as one .json file each in
/// %AppData%/Invigoration/Themes — made by duplicating any theme in Customize → Manage Themes, and
/// shareable by copying the file, the same way color schemes work (ColorSchemeLibrary).
/// </summary>
public static class ThemeLibrary
{
    public const string DefaultId = "default";
    public const string DiabloIIId = "diablo2";
    public const string StarCraftId = "starcraft";
    public const string WarcraftId = "warcraft";

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() },
    };

    private static readonly Lock SyncRoot = new();
    private static List<ChatTheme>? _custom;

    /// <summary>Raised after a custom theme is saved or deleted, so bots using it can redraw.</summary>
    public static event Action? ThemesChanged;

    /// <summary>Test hook: points the library at a scratch folder instead of the real config directory.</summary>
    public static string? DirectoryOverride { get; set; }

    public static string Directory => DirectoryOverride ?? Path.Combine(ConfigStore.DefaultConfigDirectory(), "Themes");

    public static IReadOnlyList<ChatTheme> BuiltIns { get; } =
    [
        new()
        {
            Id = DefaultId,
            Name = "Default",
            IsBuiltIn = true,
            Frame = ThemeFrameStyle.None,
            Layout = ThemeLayout.Standard,
            ColorScheme = ChatColorScheme.Invigoration,
            ShowNamePlate = false,
            Materials = new() { Band = 0x3A3A3A, BandShade = 0x1B1B1B, Trim = 0x6A6A6A, TrimShade = 0x3A3A3A, Metal = 0x8A8A8A, Accent = 0x2CACE8, Well = 0x1B1B1B },
        },
        new()
        {
            Id = DiabloIIId,
            Name = "Diablo II",
            IsBuiltIn = true,
            Frame = ThemeFrameStyle.DiabloII,
            Layout = ThemeLayout.CharacterDock,
            ColorScheme = ChatColorScheme.DiabloII,
            ShowChatGem = true,
            HeaderFont = "Georgia,Times New Roman",
            Materials = new() { Band = 0x4A4238, BandShade = 0x181511, Trim = 0xB8975A, TrimShade = 0x6B5430, Metal = 0xA8894C, Accent = 0xD8BE7A, Well = 0x050404 },
        },
        new()
        {
            Id = StarCraftId,
            Name = "StarCraft",
            IsBuiltIn = true,
            Frame = ThemeFrameStyle.StarCraft,
            Layout = ThemeLayout.Standard,
            ColorScheme = ChatColorScheme.StarCraft,
            HeaderFont = "Menlo,Consolas,Courier New",
            Materials = new() { Band = 0x4C5866, BandShade = 0x14191F, Trim = 0x8C9BA8, TrimShade = 0x39434D, Metal = 0xA9B4BE, Accent = 0x4FD8FF, Well = 0x050B12 },
        },
        new()
        {
            Id = WarcraftId,
            Name = "Warcraft",
            IsBuiltIn = true,
            Frame = ThemeFrameStyle.Warcraft,
            Layout = ThemeLayout.Standard,
            ColorScheme = ChatColorScheme.Warcraft,
            HeaderFont = "Palatino,Palatino Linotype,Book Antiqua,Georgia",
            Materials = new() { Band = 0x7A5634, BandShade = 0x2A1A0E, Trim = 0xD2AA4E, TrimShade = 0x7A5A1E, Metal = 0x62666C, Accent = 0xF0C75E, Well = 0x110B06 },
        },
    ];

    /// <summary>Built-ins first, in their fixed order, then custom themes by name.</summary>
    public static IReadOnlyList<ChatTheme> All()
    {
        lock (SyncRoot)
        {
            _custom ??= LoadCustom();
            return [.. BuiltIns, .. _custom.OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase)];
        }
    }

    /// <summary>The theme with this id, or Default when there isn't one (never set, or a custom theme that's since been deleted).</summary>
    public static ChatTheme Resolve(string? id) =>
        All().FirstOrDefault(t => string.Equals(t.Id, id, StringComparison.OrdinalIgnoreCase)) ?? BuiltIns[0];

    /// <summary>
    /// The theme id a bot actually uses. A bot saved before themes existed has no ThemeId, but may
    /// have had the old per-bot "D2 Style" switch on — that becomes the Diablo II theme, so nobody's
    /// layout changes on upgrade.
    /// </summary>
    public static string ThemeIdFor(BotConfig config) =>
        !string.IsNullOrWhiteSpace(config.ThemeId) ? config.ThemeId
        : config.UseD2ChatLayout ? DiabloIIId
        : DefaultId;

    public static ChatTheme ResolveFor(BotConfig config) => Resolve(ThemeIdFor(config));

    /// <summary>Points a bot at a theme, keeping the legacy D2 Style flag in step so an older Invigoration reading the same config still lays it out the same way.</summary>
    public static void AssignTo(BotConfig config, ChatTheme theme)
    {
        config.ThemeId = theme.Id;
        config.UseD2ChatLayout = theme.Layout == ThemeLayout.CharacterDock;
    }

    /// <summary>A new, unsaved custom theme copied from any theme (built-in or not), with its own id.</summary>
    public static ChatTheme Duplicate(ChatTheme source, string name) => new()
    {
        Id = "custom-" + Guid.NewGuid().ToString("N")[..12],
        Name = name,
        Frame = source.Frame,
        Layout = source.Layout,
        ColorScheme = source.ColorScheme,
        ShowChatGem = source.ShowChatGem,
        ShowNamePlate = source.ShowNamePlate,
        HeaderFont = source.HeaderFont,
        Materials = source.Materials.Clone(),
    };

    /// <summary>Writes a custom theme as "&lt;id&gt;.json", replacing any earlier version of it.</summary>
    /// <exception cref="InvalidOperationException">The theme is built in.</exception>
    public static void Save(ChatTheme theme)
    {
        if (theme.IsBuiltIn || BuiltIns.Any(b => string.Equals(b.Id, theme.Id, StringComparison.OrdinalIgnoreCase)))
        {
            throw new InvalidOperationException($"\"{theme.Name}\" is a built-in theme and can't be overwritten — duplicate it instead.");
        }

        lock (SyncRoot)
        {
            System.IO.Directory.CreateDirectory(Directory);
            File.WriteAllText(PathFor(theme.Id), JsonSerializer.Serialize(theme, JsonOptions));
            _custom = null;
        }

        ThemesChanged?.Invoke();
    }

    /// <summary>Removes a custom theme. Bots still pointing at it fall back to Default (see Resolve). Built-in ids are ignored.</summary>
    public static void Delete(string id)
    {
        if (BuiltIns.Any(b => string.Equals(b.Id, id, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        lock (SyncRoot)
        {
            var path = PathFor(id);
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            _custom = null;
        }

        ThemesChanged?.Invoke();
    }

    /// <summary>Test hook: forgets the loaded custom themes so the next read goes back to disk.</summary>
    public static void ResetCacheForTests()
    {
        lock (SyncRoot)
        {
            _custom = null;
        }
    }

    private static string PathFor(string id)
    {
        var safe = string.Concat(id.Where(c => char.IsLetterOrDigit(c) || c is '-' or '_'));
        return Path.Combine(Directory, (safe.Length > 0 ? safe : "theme") + ".json");
    }

    private static List<ChatTheme> LoadCustom()
    {
        if (!System.IO.Directory.Exists(Directory))
        {
            return [];
        }

        var themes = new List<ChatTheme>();
        foreach (var file in System.IO.Directory.GetFiles(Directory, "*.json"))
        {
            try
            {
                if (JsonSerializer.Deserialize<ChatTheme>(File.ReadAllText(file), JsonOptions) is { Id.Length: > 0 } theme &&
                    !BuiltIns.Any(b => string.Equals(b.Id, theme.Id, StringComparison.OrdinalIgnoreCase)))
                {
                    themes.Add(theme);
                }
            }
            catch (Exception ex) when (ex is JsonException or IOException)
            {
                // A hand-edited or half-copied file — skip it rather than lose every other theme.
            }
        }

        return themes;
    }
}
