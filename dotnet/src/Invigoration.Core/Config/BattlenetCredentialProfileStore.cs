using System.Text.Json;

namespace Invigoration.Core.Config;

/// <summary>
/// The shared (cross-bot) list of named Battle.net logins — persisted at
/// %AppData%/Invigoration/battlenet-credential-profiles.json. Each profile's
/// actual signed-in session is cached by Stimpak (the native library behind
/// the SC2/SC:R/WC3:R connections) to a file keyed by the profile's stable
/// Id, under %AppData%/Invigoration/BattlenetCredentials/&lt;id&gt;.bin — see
/// CredentialFilePath. Deliberately not product-namespaced: the whole point
/// is that an SC2 bot and a future WC3:Reforged bot on the same Battle.net
/// account can reference the same profile and share that one file.
/// </summary>
public static class BattlenetCredentialProfileStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private static readonly Lock SyncRoot = new();
    private static List<BattlenetCredentialProfile>? _cache;
    private static string? _configDirectoryOverride;

    /// <summary>
    /// Test-only hook that redirects every read/write to an isolated directory instead of the
    /// real %AppData%/Invigoration — see Invigoration.Core.Tests'
    /// BattlenetCredentialProfileStoreFixture. Never set this outside a test. Setting it also
    /// drops the in-memory cache so the next Profiles access re-reads from the new location.
    /// </summary>
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

    public static string FilePath => Path.Combine(ConfigDirectory, "battlenet-credential-profiles.json");

    public static List<BattlenetCredentialProfile> Profiles => _cache ??= LoadFromDisk();

    /// <summary>Raised after every Save() — lets an open Manage Battle.net Profiles window or a bot's Config window picker refresh.</summary>
    public static event Action? ProfilesChanged;

    public static BattlenetCredentialProfile? Find(string id) =>
        string.IsNullOrEmpty(id) ? null : Profiles.FirstOrDefault(p => p.Id == id);

    /// <summary>Creates a new profile and persists it immediately — used both by the Manage Profiles window's "Add" and by BotEngine.Sc2.cs's auto-create-on-first-use fallback, neither of which has a separate explicit "Save" step of its own.</summary>
    public static BattlenetCredentialProfile CreateAndSave(string name)
    {
        var profile = new BattlenetCredentialProfile { Name = string.IsNullOrWhiteSpace(name) ? "New Profile" : name.Trim() };
        Profiles.Add(profile);
        Save();
        return profile;
    }

    /// <summary>Removes the profile entry and its cached credential file, if any. Callers are responsible for warning the user first if any bot config still references this Id — the store itself has no visibility into bots.json.</summary>
    public static void Delete(string id)
    {
        Profiles.RemoveAll(p => p.Id == id);
        Save();

        var credentialPath = CredentialFilePath(id);
        if (File.Exists(credentialPath))
        {
            try
            {
                File.Delete(credentialPath);
            }
            catch (IOException)
            {
                // e.g. a live bot still has the file open — leave it, a harmless orphan.
            }
        }
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

    /// <summary>Where this profile's Stimpak session caches its signed-in credential.</summary>
    public static string CredentialFilePath(string profileId) =>
        Path.Combine(ConfigDirectory, "BattlenetCredentials", profileId + ".bin");

    /// <summary>Cheap local "is this signed in" check — file existence/non-empty, not a network round-trip.</summary>
    public static bool HasCachedCredential(string profileId)
    {
        var path = CredentialFilePath(profileId);
        return File.Exists(path) && new FileInfo(path).Length > 0;
    }

    /// <summary>
    /// Records the real signed-in Battle.net username against a profile — called whenever an
    /// AccountConnected SC2Event arrives (BotEngine.Sc2.cs's live connect path, and the
    /// throwaway client BattlenetCredentialProfilesViewModel's standalone Sign In uses). A
    /// no-op if the profile's already gone (e.g. deleted mid-connect) or the tag hasn't
    /// actually changed, so this can be called on every AccountConnected without spamming
    /// ProfilesChanged/disk writes on an already-known account.
    /// </summary>
    public static void UpdateBattleTag(string profileId, string? battleTag)
    {
        if (string.IsNullOrEmpty(battleTag))
        {
            return;
        }

        var profile = Find(profileId);
        if (profile is null || profile.BattleTag == battleTag)
        {
            return;
        }

        profile.BattleTag = battleTag;
        Save();
    }

    private static List<BattlenetCredentialProfile> LoadFromDisk()
    {
        if (!File.Exists(FilePath))
        {
            return [];
        }

        var loaded = JsonSerializer.Deserialize<List<BattlenetCredentialProfile>>(File.ReadAllText(FilePath), JsonOptions);
        return loaded ?? [];
    }
}
