using System.Text.Json;

namespace Invigoration.Core.Clan;

/// <summary>
/// A shared (cross-bot) roster of clan members, persisted at
/// %AppData%/Invigoration/clan-members.json — one roster for the whole
/// install, not per-bot, since the same clan structure is useful across
/// every bot the user runs. Cached in memory after first load so every
/// connected bot and the management window see the same live list; call
/// <see cref="Save"/> after any edit to persist it.
/// </summary>
public static class ClanRosterStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private static readonly Lock SyncRoot = new();
    private static List<ClanMember>? _cache;

    public static string FilePath => Path.Combine(Config.ConfigStore.DefaultConfigDirectory(), "clan-members.json");

    public static List<ClanMember> Members => _cache ??= LoadFromDisk();

    /// <summary>Raised after every Save() — lets an open bot tab or management window pick up roster changes made elsewhere (a chat command, another window) without needing to reopen.</summary>
    public static event Action? RosterChanged;

    /// <summary>Finds the member whose primary name or an alias matches the given Battle.net username, or null if untracked. Unscoped by server — for editing/management lookups, not authorization (use FindTrusted for those).</summary>
    public static ClanMember? Find(string username) => Members.FirstOrDefault(m => m.Matches(username));

    /// <summary>
    /// Server-scoped lookup for authorization decisions (bot-master check,
    /// rank-based permission grants, ban checks) — see
    /// <see cref="ClanMember.MatchesOnServer"/> for why this exists instead
    /// of just using <see cref="Find"/> everywhere.
    /// </summary>
    public static ClanMember? FindTrusted(string username, string speakerServer) =>
        Members.FirstOrDefault(m => m.MatchesOnServer(username, speakerServer));

    /// <summary>
    /// Stamps a tracked member's LastSeenUtc as now, and — when known —
    /// LastSeenProduct/LastSeenServer. If they're untracked and
    /// <paramref name="defaultRankIfNew"/> is non-empty, creates a new
    /// (auto-tracked, not a formal clan member — see ClanMember.IsClanMember)
    /// roster entry for them with that rank first — this is what builds an
    /// ongoing roster of everyone who's ever spoken, not just people
    /// explicitly added, so ranks (including a "banned" one) can be handed
    /// out later. Passing null/empty for <paramref name="defaultRankIfNew"/>
    /// keeps the old behavior: a no-op for anyone not already tracked.
    /// </summary>
    public static void RecordSeen(string username, string? defaultRankIfNew = null, string? product = null, string? server = null)
    {
        // Locked end-to-end (not just the file write) so two bots' chat
        // handlers racing to auto-register the same first-time username
        // can't both miss the Find() and create duplicate entries.
        lock (SyncRoot)
        {
            var member = Find(username);
            if (member is null)
            {
                if (string.IsNullOrWhiteSpace(defaultRankIfNew))
                {
                    return;
                }

                member = new ClanMember { Name = username, Rank = defaultRankIfNew, IsClanMember = false };
                Members.Add(member);
            }

            member.LastSeenUtc = DateTime.UtcNow;
            if (!string.IsNullOrEmpty(product))
            {
                member.LastSeenProduct = product;
            }

            if (!string.IsNullOrEmpty(server))
            {
                member.LastSeenServer = server;
            }

            SaveLocked();
        }
    }

    /// <summary>
    /// Opportunistically updates an already-tracked member's last-seen
    /// game/server from a presence sighting (joining/showing up in a
    /// channel) rather than actual chat — a no-op for anyone not already
    /// tracked, since presence alone shouldn't auto-create a roster entry
    /// the way actually talking does (see RecordSeen).
    /// </summary>
    public static void RecordProductSeen(string username, string product, string server)
    {
        lock (SyncRoot)
        {
            var member = Find(username);
            if (member is null)
            {
                return;
            }

            member.LastSeenProduct = product;
            member.LastSeenServer = server;
            SaveLocked();
        }
    }

    /// <summary>Locked so concurrent Save() calls from multiple bots (or a bot and the Clan Members window) can't collide writing the same file.</summary>
    public static void Save()
    {
        lock (SyncRoot)
        {
            SaveLocked();
        }
    }

    private static void SaveLocked()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        File.WriteAllText(FilePath, JsonSerializer.Serialize(Members, JsonOptions));
        RosterChanged?.Invoke();
    }

    private static List<ClanMember> LoadFromDisk()
    {
        if (!File.Exists(FilePath))
        {
            return [];
        }

        return JsonSerializer.Deserialize<List<ClanMember>>(File.ReadAllText(FilePath), JsonOptions) ?? [];
    }
}
