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
            // First run only — keyed on the file not existing rather than the list being empty, so
            // someone who deliberately deletes every profile doesn't get these back on next launch.
            var seeded = DefaultProfiles();
            TrySaveSeed(seeded);
            return seeded;
        }

        var loaded = JsonSerializer.Deserialize<List<HotlineServerProfile>>(File.ReadAllText(FilePath), JsonOptions);
        return loaded ?? [];
    }

    private static void TrySaveSeed(List<HotlineServerProfile> seeded)
    {
        try
        {
            Directory.CreateDirectory(ConfigDirectory);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(seeded, JsonOptions));
        }
        catch (IOException)
        {
            // The in-memory copy still has them for this run; nothing here is worth failing over.
        }
    }

    /// <summary>
    /// A couple of long-running public servers to start from, so a new Hotline tab isn't an empty
    /// box asking for an IP address. Both are real servers this client has been tested against.
    ///
    /// Deliberately NOT included, whatever any one person's own setup looks like:
    ///   - Auto-connect. Nothing dials out on startup unless someone asks it to.
    ///   - Any login or password. These connect as a guest; the nickname is "Guest".
    ///   - Auto-accepting the server's agreement. Agreeing to someone's rules on their behalf
    ///     isn't a default to ship (see HotlineTransactionClient.AutoAcceptAgreement).
    ///
    /// The Discord relay names ARE included: each of these servers bridges its chat to Discord
    /// under a specific account name, and without it the relay's messages show up as an ordinary
    /// user talking rather than as relayed Discord chat.
    /// </summary>
    public static List<HotlineServerProfile> DefaultProfiles() =>
    [
        new()
        {
            Name = "MacDomain",
            Host = "62.116.228.143",
            Port = 5500,
            DiscordRelayUsername = "Discord",
        },
        new()
        {
            Name = "HL Central",
            Host = "74.208.191.206",
            Port = 5500,
            DiscordRelayUsername = "Relay",
            DiscordRelayPrefix = "Discord |",

            // This one speaks HOPE, so the password is MAC'd rather than sent obfuscated if
            // anyone does put credentials on it later. Falls back on its own if it ever stops.
            UseSecureLogin = true,
        },
    ];
}
