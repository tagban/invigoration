using System.Text.Json;

namespace Invigoration.Core.Tracking;

/// <summary>
/// Last-seen/trivia-score tracking for non-Battle.net protocols (Hotline today) — persisted at
/// %AppData%/Invigoration/protocol-users.json, deliberately its own file and its own in-memory list
/// rather than folded into Clan.ClanRosterStore's clan-members.json, per explicit request: these
/// users were never added to a clan roster and must not show up in any "CLAN" list or command.
/// Unlike ClanRosterStore (which only scores/tracks a name someone explicitly "clanadd"ed),
/// entries here are upserted automatically just from being seen chatting — there's no equivalent
/// opt-in step for a Hotline server, and the whole point per the original request is passive
/// last-seen/score tracking.
/// </summary>
public static class ProtocolUserTrackingStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private static readonly Lock SyncRoot = new();
    private static List<TrackedUser>? _cache;

    public static string FilePath => Path.Combine(Config.ConfigStore.DefaultConfigDirectory(), "protocol-users.json");

    public static List<TrackedUser> Users => _cache ??= LoadFromDisk();

    public static TrackedUser? Find(string name, string protocol, string server) =>
        Users.FirstOrDefault(u => u.Matches(name, protocol, server));

    /// <summary>Marks a user as seen just now — creates a new entry on first sight, otherwise just bumps LastSeenUtc. Called for every chat line seen from a real (non-ghost) user, not just command use.</summary>
    public static TrackedUser MarkSeen(string name, string protocol, string server)
    {
        lock (SyncRoot)
        {
            var user = Find(name, protocol, server);
            if (user is null)
            {
                user = new TrackedUser { Name = name, Protocol = protocol, Server = server };
                Users.Add(user);
            }

            user.LastSeenUtc = DateTime.UtcNow;
            Save();
            return user;
        }
    }

    /// <summary>Adds to a user's trivia score (creating/marking them seen first if this is their first score), returning the new total.</summary>
    public static double AddScore(string name, string protocol, string server, double delta)
    {
        var user = MarkSeen(name, protocol, server);
        lock (SyncRoot)
        {
            user.TriviaScore += delta;
            Save();
            return user.TriviaScore;
        }
    }

    /// <summary>Top scorers for one protocol+server (e.g. "Hotline"/"bigredh.com:5500"), highest first — scoped to a single server, not every server under the protocol, matching that a trivia round never crosses server boundaries (see HotlineTriviaHost's remarks). Mirrors ClanRosterStore-backed BotEngine.HandleTriviaScoreAsync's own leaderboard shape.</summary>
    public static List<TrackedUser> GetLeaderboard(string protocol, string server, int take = 10) =>
        Users.Where(u => u.Protocol.Equals(protocol, StringComparison.OrdinalIgnoreCase) &&
                u.Server.Equals(server, StringComparison.OrdinalIgnoreCase) && u.TriviaScore != 0)
            .OrderByDescending(u => u.TriviaScore)
            .Take(take)
            .ToList();

    private static void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        File.WriteAllText(FilePath, JsonSerializer.Serialize(Users, JsonOptions));
    }

    private static List<TrackedUser> LoadFromDisk()
    {
        if (!File.Exists(FilePath))
        {
            return [];
        }

        var loaded = JsonSerializer.Deserialize<List<TrackedUser>>(File.ReadAllText(FilePath), JsonOptions);
        return loaded ?? [];
    }
}
