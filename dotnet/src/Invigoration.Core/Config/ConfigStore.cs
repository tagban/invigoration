using System.Text.Json;
using System.Text.Json.Serialization;

namespace Invigoration.Core.Config;

/// <summary>
/// Loads/saves the list of configured bots as JSON. Replaces modFunctions.bas's
/// GetStuff/WriteStuff (Win32 GetPrivateProfileString/WritePrivateProfileString
/// INI helpers) with a cross-platform format and multi-bot support the
/// original's single-INI-section design didn't have.
/// </summary>
public sealed class ConfigStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() },
    };

    public string FilePath { get; }

    public ConfigStore(string? filePath = null)
    {
        FilePath = filePath ?? Path.Combine(DefaultConfigDirectory(), "bots.json");
    }

    public static string DefaultConfigDirectory()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        return Path.Combine(appData, "Invigoration");
    }

    public List<BotConfig> Load()
    {
        if (!File.Exists(FilePath))
        {
            return [];
        }

        var json = File.ReadAllText(FilePath);
        return JsonSerializer.Deserialize<List<BotConfig>>(json, JsonOptions) ?? [];
    }

    public void Save(IReadOnlyList<BotConfig> bots)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        var json = JsonSerializer.Serialize(bots, JsonOptions);
        File.WriteAllText(FilePath, json);
    }
}
