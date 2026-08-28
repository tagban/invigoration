using System.Text.Json;
using Invigoration.Core.Config;

namespace Invigoration.Core.Hotline;

/// <summary>
/// The saved list of Hotline server profiles — persisted at
/// %AppData%/Invigoration/hotline-server-profiles.json, same shape as
/// BattlenetCredentialProfileStore (including its test-only ConfigDirectoryOverride hook). Global,
/// not per-bot: Hotline is its own protocol/tab-group, not something individual Battle.net bots
/// have one each of.
/// </summary>
public static class HotlineServerProfileStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private static readonly Lock SyncRoot = new();
    private static List<HotlineServerProfile>? _cache;
    private static string? _configDirectoryOverride;

    /// <summary>Test-only hook — see BattlenetCredentialProfileStore.ConfigDirectoryOverride's remarks, same reasoning applies here.</summary>
    public static string? ConfigDirectoryOverride
    {
        get => _configDirectoryOverride;
        set
        {
            _configDirectoryOverride = value;
            _cache = null;
        }
    }

    private static string ConfigDirectory => ConfigDirectoryOverride ?? ConfigStore.DefaultConfigDirectory();

    public static string FilePath => Path.Combine(ConfigDirectory, "hotline-server-profiles.json");

    public static List<HotlineServerProfile> Profiles => _cache ??= LoadFromDisk();

    public static event Action? ProfilesChanged;

    public static HotlineServerProfile? Find(string id) =>
        string.IsNullOrEmpty(id) ? null : Profiles.FirstOrDefault(p => p.Id == id);

    public static HotlineServerProfile CreateAndSave(string name, string host, ushort port)
    {
        var profile = new HotlineServerProfile
        {
            Name = string.IsNullOrWhiteSpace(name) ? "New Server" : name.Trim(),
            Host = host,
            Port = port,
        };
        Profiles.Add(profile);
        Save();
        return profile;
    }

    public static void Delete(string id)
    {
        Profiles.RemoveAll(p => p.Id == id);
        Save();
    }

    public static void Save()
    {
        lock (SyncRoot)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(Profiles, JsonOptions));
            ProfilesChanged?.Invoke();
        }
    }

    private static List<HotlineServerProfile> LoadFromDisk()
    {
        if (!File.Exists(FilePath))
        {
            return [];
        }

        var loaded = JsonSerializer.Deserialize<List<HotlineServerProfile>>(File.ReadAllText(FilePath), JsonOptions);
        return loaded ?? [];
    }
}
