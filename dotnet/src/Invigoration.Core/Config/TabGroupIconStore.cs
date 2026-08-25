using System.Text.Json;

namespace Invigoration.Core.Config;

/// <summary>
/// Which icon key (see GameIconLoader/IconOverrideStore) a named bot tab group
/// (BotConfig.TabGroup) shows on its collapsed top-level tab — persisted at
/// %AppData%/Invigoration/tab-group-icons.json, keyed by the group name itself
/// (groups aren't a separate persisted entity anywhere else, so the name is
/// the only stable key available). With no entry, BotGroupTabViewModel falls
/// back to borrowing its first member bot's own icon.
/// </summary>
public static class TabGroupIconStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private static readonly Lock SyncRoot = new();
    private static Dictionary<string, string>? _cache;
    private static string? _configDirectoryOverride;

    /// <summary>Test-only hook, same pattern as BattlenetCredentialProfileStore's — redirects reads/writes to an isolated directory instead of the real %AppData%/Invigoration.</summary>
    public static string? ConfigDirectoryOverride
    {
        get => _configDirectoryOverride;
        set
        {
            _configDirectoryOverride = value;
            _cache = null;
        }
    }

    private static string FilePath => Path.Combine(_configDirectoryOverride ?? ConfigStore.DefaultConfigDirectory(), "tab-group-icons.json");

    private static Dictionary<string, string> Assignments => _cache ??= LoadFromDisk();

    public static string? GetIconKey(string groupName) =>
        !string.IsNullOrEmpty(groupName) && Assignments.TryGetValue(groupName, out var key) ? key : null;

    public static void SetIconKey(string groupName, string? iconKey)
    {
        if (string.IsNullOrEmpty(groupName))
        {
            return;
        }

        if (string.IsNullOrEmpty(iconKey))
        {
            Assignments.Remove(groupName);
        }
        else
        {
            Assignments[groupName] = iconKey;
        }

        Save();
    }

    private static void Save()
    {
        lock (SyncRoot)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(Assignments, JsonOptions));
        }
    }

    private static Dictionary<string, string> LoadFromDisk()
    {
        if (!File.Exists(FilePath))
        {
            return [];
        }

        var loaded = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(FilePath), JsonOptions);
        return loaded ?? [];
    }
}
