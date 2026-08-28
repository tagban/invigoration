using System.Text.Json;

namespace Invigoration.Core.Clan;

/// <summary>
/// The shared (cross-bot) list of predefined ranks — persisted at
/// %AppData%/Invigoration/clan-ranks.json. A controlled list (rather than
/// free-text, which risks typos creating phantom ranks nothing actually
/// grants access to, or that no bot recognizes for its auto-behaviors) is
/// what lets the Seen List's Rank field be a dropdown, gives every bot's
/// PermissionLevel.Ranks a shared vocabulary to reference, and gives each
/// rank's auto-whisper/auto-kick/auto-ban behavior somewhere to live.
/// Seeded with a few sensible defaults on first use, matching BotConfig's
/// own DefaultRank ("Trivia Participant") and BannedRank ("Banned") —
/// "Banned" ships with AutoKick on, since banning someone and not also
/// removing them from the channel isn't usually the intent. AllowedCommands
/// on each rank is what used to live on the now-removed BotConfig.PermissionLevels
/// — command access is granted directly on the rank instead of a separate list.
/// </summary>
public static class ClanRankStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private static readonly Lock SyncRoot = new();
    private static List<ClanRank>? _cache;

    public static string FilePath => Path.Combine(Config.ConfigStore.DefaultConfigDirectory(), "clan-ranks.json");

    public static List<ClanRank> Ranks => _cache ??= LoadFromDisk();

    /// <summary>Raised after every Save() — lets an open Seen List window pick up rank-list changes made elsewhere without needing to reopen.</summary>
    public static event Action? RanksChanged;

    /// <summary>Finds a predefined rank by name (case-insensitive), or null if it isn't one of the controlled ranks (e.g. a legacy free-text value from before this existed).</summary>
    public static ClanRank? Find(string name) =>
        string.IsNullOrEmpty(name) ? null : Ranks.FirstOrDefault(r => r.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

    public static void Save()
    {
        lock (SyncRoot)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(Ranks, JsonOptions));
            RanksChanged?.Invoke();
        }
    }

    private static List<ClanRank> LoadFromDisk()
    {
        if (!File.Exists(FilePath))
        {
            return DefaultRanks();
        }

        var loaded = JsonSerializer.Deserialize<List<ClanRank>>(File.ReadAllText(FilePath), JsonOptions);
        return loaded is { Count: > 0 } ? loaded : DefaultRanks();
    }

    private static List<ClanRank> DefaultRanks() =>
    [
        new ClanRank
        {
            Name = "Leader",
            AllowedCommands =
            [
                "kick", "ban", "join", "home", "clanadd", "clanremove", "clanrank", "clanalias",
                "clanunalias", "claninfo", "clanlist", "clanscore", "trivia", "idle", "disconnect", "reconnect",
                "user", "useroff", "trigger", "sysinfo", "ver", "uptime", "about",
                "bancount", "kickcount", "joincount",
            ],
        },
        new ClanRank
        {
            Name = "Officer",
            AllowedCommands =
            [
                "kick", "ban", "join", "home", "clanadd", "clanrank", "clanalias", "clanunalias",
                "claninfo", "clanlist", "clanscore", "trivia", "user", "useroff", "sysinfo", "ver",
                "uptime", "about", "bancount", "kickcount", "joincount",
            ],
        },
        new ClanRank
        {
            Name = "Member",
            AllowedCommands = ["claninfo", "clanlist", "sysinfo", "ver", "uptime", "about"],
        },
        // Trivia score/join are already unconditionally open to everyone (see BotEngine.Commands.IsAuthorized),
        // so this rank needs no AllowedCommands of its own — it exists mainly so auto-tracked chatters have
        // somewhere to land and show up on the trivia leaderboard.
        new ClanRank { Name = "Trivia Participant" },
        new ClanRank { Name = "Banned", AutoKick = true },
    ];
}
