using System.Text.Json;
using Invigoration.Core.Chat;

namespace Invigoration.Core.Config;

/// <summary>
/// A folder of shareable color-scheme .json files (NamedCustomPalette) at
/// %AppData%/Invigoration/Colors — populated on first run with the three
/// built-in schemes (both as a starting point for building a custom one and
/// as a worked example of the file format), and added to by "Save to
/// Library" or "Import..." in the config window's Custom Colors section.
/// Dropping a file here by hand (e.g. one a friend emailed) makes it show
/// up in the library too, since it's just scanned by filename.
/// </summary>
public static class ColorSchemeLibrary
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public static string Directory => Path.Combine(ConfigStore.DefaultConfigDirectory(), "Colors");

    /// <summary>Creates the Colors folder and seeds it with the built-in schemes, but only the first time — never overwrites what's there once the folder exists, so hand edits/deletions stick.</summary>
    public static void EnsureBuiltInSchemesExist()
    {
        if (System.IO.Directory.Exists(Directory))
        {
            return;
        }

        System.IO.Directory.CreateDirectory(Directory);
        Save(new NamedCustomPalette { Name = "Invigoration", Colors = ChatPalette.Invigoration.ToCustom() });
        Save(new NamedCustomPalette { Name = "BNU`Bot StarCraft", Colors = ChatPalette.StarCraft.ToCustom() });
        Save(new NamedCustomPalette { Name = "BNU`Bot Diablo", Colors = ChatPalette.DiabloII.ToCustom() });
    }

    /// <summary>All valid scheme files currently in the library, sorted by name.</summary>
    public static IReadOnlyList<(string FilePath, string Name)> ListSchemes()
    {
        if (!System.IO.Directory.Exists(Directory))
        {
            return [];
        }

        var results = new List<(string FilePath, string Name)>();
        foreach (var file in System.IO.Directory.GetFiles(Directory, "*.json"))
        {
            try
            {
                results.Add((file, Load(file).Name));
            }
            catch (JsonException)
            {
                // Not a valid scheme file — skip it rather than fail the whole listing.
            }
        }

        return results.OrderBy(r => r.Name, StringComparer.OrdinalIgnoreCase).ToList();
    }

    public static NamedCustomPalette Load(string filePath) =>
        JsonSerializer.Deserialize<NamedCustomPalette>(File.ReadAllText(filePath), JsonOptions)
        ?? throw new JsonException($"'{filePath}' is not a valid color scheme file.");

    /// <summary>Writes a scheme as "&lt;Name&gt;.json", overwriting any existing file of that name. Returns the path written.</summary>
    public static string Save(NamedCustomPalette scheme)
    {
        System.IO.Directory.CreateDirectory(Directory);
        var path = Path.Combine(Directory, SanitizeFileName(scheme.Name) + ".json");
        File.WriteAllText(path, JsonSerializer.Serialize(scheme, JsonOptions));
        return path;
    }

    private static string SanitizeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(name.Select(c => invalid.Contains(c) ? '_' : c).ToArray()).Trim();
        return cleaned.Length == 0 ? "Custom Scheme" : cleaned;
    }
}
