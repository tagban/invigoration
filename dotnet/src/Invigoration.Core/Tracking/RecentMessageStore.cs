using System.Text.Json;

namespace Invigoration.Core.Tracking;

/// <summary>
/// A small local ring-buffer of the last ~10 chat lines per protocol+server, persisted at
/// %AppData%/Invigoration/recent-messages.json — a client-side convenience so reconnecting to a
/// server shows roughly where the conversation left off, for any protocol/server with no
/// server-side memory of its own. Per explicit request, only Hotline populates this today, but the
/// mechanism itself is protocol/server-keyed the same way Tracking.ProtocolUserTrackingStore is,
/// so IRC/FFXI can reuse it later without a new store.
///
/// This is deliberately NOT the same thing as Hotline's own "2.5" server-side chat history
/// extension (HotlineTransactionClient.GetChatHistoryAsync) — a server that speaks that extension
/// already sends a real dump of its own persisted history on request, which is authoritative and
/// can go back further than 10 lines. This local cache exists specifically for everything else: a
/// pre-2.5 (1.2.3+) server that has no memory of the conversation at all. See
/// HotlineSessionViewModel's own use of both — it shows the server's real history when available,
/// and only falls back to this local cache when it isn't.
/// </summary>
public static class RecentMessageStore
{
    public const int RetentionCount = 10;

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private static readonly Lock SyncRoot = new();
    private static Dictionary<string, List<RecentMessage>>? _cache;

    public static string FilePath => Path.Combine(Config.ConfigStore.DefaultConfigDirectory(), "recent-messages.json");

    private static Dictionary<string, List<RecentMessage>> Store => _cache ??= LoadFromDisk();

    private static string Key(string protocol, string server) => $"{protocol}:{server}";

    /// <summary>The last (up to RetentionCount) cached lines for one protocol+server, oldest-first.</summary>
    public static IReadOnlyList<RecentMessage> GetRecent(string protocol, string server) =>
        Store.TryGetValue(Key(protocol, server), out var list) ? list : [];

    /// <summary>Appends one line, trimming the oldest once past RetentionCount.</summary>
    public static void Append(string protocol, string server, string text, DateTimeOffset? timestampUtc = null)
    {
        lock (SyncRoot)
        {
            var key = Key(protocol, server);
            if (!Store.TryGetValue(key, out var list))
            {
                list = [];
                Store[key] = list;
            }

            list.Add(new RecentMessage { Text = text, TimestampUtc = timestampUtc });
            while (list.Count > RetentionCount)
            {
                list.RemoveAt(0);
            }

            Save();
        }
    }

    private static void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        File.WriteAllText(FilePath, JsonSerializer.Serialize(Store, JsonOptions));
    }

    private static Dictionary<string, List<RecentMessage>> LoadFromDisk()
    {
        if (!File.Exists(FilePath))
        {
            return [];
        }

        var loaded = JsonSerializer.Deserialize<Dictionary<string, List<RecentMessage>>>(File.ReadAllText(FilePath), JsonOptions);
        return loaded ?? [];
    }
}
