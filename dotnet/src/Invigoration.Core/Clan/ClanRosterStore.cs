using System.Text.Json;
using Invigoration.Core.Chat;
using Invigoration.Core.Protocol;

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

    /// <summary>How long a rapid run of auto-tracking mutations (RecordSeen/RecordProductSeen) can pile up before the next actual disk write — see MarkDirtyAndNotify's remarks.</summary>
    private static readonly TimeSpan SaveDebounceInterval = TimeSpan.FromSeconds(2);
    private static Timer? _saveTimer;
    private static bool _saveDirty;

    public static string FilePath => Path.Combine(Config.ConfigStore.DefaultConfigDirectory(), "clan-members.json");

    /// <summary>
    /// Forces the next Find/FindTrusted to rebuild NameIndex immediately rather than waiting out
    /// NameIndexTtl — needed by any caller (tests, mainly) that mutates Members directly
    /// (Add/RemoveAll/in-place edit) instead of through RecordSeen/RecordProductSeen/Save, which
    /// already invalidate on their own. Without this, a test that adds a member and immediately
    /// looks it up can race a *different* test's still-warm cache from within the same TTL
    /// window — confirmed live as real failures, not a hypothetical.
    /// </summary>
    public static void InvalidateNameIndex() => _nameIndex = null;

    public static List<ClanMember> Members => _cache ??= LoadFromDisk();

    /// <summary>Raised after every Save() — lets an open bot tab or management window pick up roster changes made elsewhere (a chat command, another window) without needing to reopen.</summary>
    public static event Action? RosterChanged;

    // Name/alias -> candidate members, keyed on the bare (server-qualifier stripped, '*'-stripped)
    // name so a single lookup handles both Find and FindTrusted's "any member whose Name or an
    // Alias matches" shape. Every Join/ShowUser/UserFlags event calls Find and/or FindTrusted
    // (RecordProductSeen, ApplyRankBehaviorsAsync) — a plain Members.FirstOrDefault(m =>
    // m.Matches(...)) scan is O(roster size) per lookup, which is fine for a roster of a few
    // dozen formal clan members but became a real bottleneck once auto-tracking (RecordSeen,
    // "everyone who's ever spoken") had grown the roster to ~1000 entries under repeated load
    // testing — confirmed live as a cause of lag reappearing at that scale even after the
    // per-event send-pileup fixes.
    //
    // Rebuilt on a short TTL rather than incrementally maintained or invalidated by comparing
    // Members.Count: Members is a plain public List<ClanMember> that callers all over the
    // codebase (including 40+ existing test call sites and the Clan Members management window)
    // mutate directly via Add/RemoveAll/in-place edits, with no CollectionChanged-style hook to
    // key an index off. A first attempt compared Members.Count to detect staleness, but that's
    // unsound: many tests each add exactly one member then remove it in a finally block, so
    // Count oscillates back to the same values across unrelated tests — confirmed live as 11 real
    // test failures (a stale index served a *different* member than the one just added, because
    // the count happened to match a previous build). A TTL sidesteps needing to detect mutations
    // at all: during an actual burst the roster is essentially static, so a sub-second staleness
    // window costs at most a couple of extra O(roster) rebuilds over the whole burst while still
    // eliminating the O(roster) cost from every individual lookup. RecordSeen's own Members.Add
    // (adding a first-time talker) and Save() (the only path an in-place rename reaches) both
    // still invalidate immediately too, so the common paths never wait out the TTL at all.
    private static readonly TimeSpan NameIndexTtl = TimeSpan.FromMilliseconds(500);
    private static Dictionary<string, List<ClanMember>>? _nameIndex;
    private static DateTime _nameIndexBuiltAtUtc = DateTime.MinValue;

    private static Dictionary<string, List<ClanMember>> NameIndex
    {
        get
        {
            if (_nameIndex is null || DateTime.UtcNow - _nameIndexBuiltAtUtc > NameIndexTtl)
            {
                _nameIndex = BuildNameIndex();
                _nameIndexBuiltAtUtc = DateTime.UtcNow;
            }

            return _nameIndex;
        }
    }

    private static Dictionary<string, List<ClanMember>> BuildNameIndex()
    {
        var index = new Dictionary<string, List<ClanMember>>(StringComparer.OrdinalIgnoreCase);
        foreach (var member in Members)
        {
            IndexKey(index, member.Name, member);
            foreach (var alias in member.Aliases)
            {
                IndexKey(index, alias, member);
            }
        }

        return index;
    }

    private static void IndexKey(Dictionary<string, List<ClanMember>> index, string entry, ClanMember member)
    {
        var key = BnetUsername.Normalize(BnetUsername.SplitServerQualifier(entry).Name);
        if (!index.TryGetValue(key, out var candidates))
        {
            candidates = [];
            index[key] = candidates;
        }

        candidates.Add(member);
    }

    private static IEnumerable<ClanMember> Candidates(string username) =>
        NameIndex.TryGetValue(BnetUsername.Normalize(username), out var candidates) ? candidates : [];

    /// <summary>Finds the member whose primary name or an alias matches the given Battle.net username, or null if untracked. Unscoped by server — for editing/management lookups, not authorization (use FindTrusted for those).</summary>
    public static ClanMember? Find(string username) => Candidates(username).FirstOrDefault(m => m.Matches(username));

    /// <summary>
    /// Server-scoped lookup for authorization decisions (bot-master check,
    /// rank-based permission grants, ban checks) — see
    /// <see cref="ClanMember.MatchesOnServer"/> for why this exists instead
    /// of just using <see cref="Find"/> everywhere.
    /// </summary>
    public static ClanMember? FindTrusted(string username, string speakerServer) =>
        Candidates(username).FirstOrDefault(m => m.MatchesOnServer(username, speakerServer));

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
                _nameIndex = null; // a first-time talker must be immediately findable, not wait out NameIndexTtl
            }

            member.LastSeenUtc = DateTime.UtcNow;
            if (!string.IsNullOrEmpty(product))
            {
                member.LastSeenProduct = product;
                RecordPlatform(member, product);
            }

            if (!string.IsNullOrEmpty(server))
            {
                member.LastSeenServer = server;
            }

            MarkDirtyAndNotify();
        }
    }

    /// <summary>
    /// Adds this product's display name (the same text LastSeenGameText already shows for "Last
    /// game") to the member's Platforms list, if it isn't already there — confirmed the field was
    /// otherwise stuck purely manual (never touched by any auto-tracking) despite LastSeenProduct
    /// already being recorded automatically. Only for classic BNCS products, the only ones that
    /// ever call RecordSeen/RecordProductSeen with a product code (see BotEngine.Bncs.cs) — SC2/
    /// SC:R/WC3:R don't currently report a product here at all, a separate gap.
    /// </summary>
    private static void RecordPlatform(ClanMember member, string product)
    {
        var displayName = BncsProduct.GetDisplayName(product);
        if (!member.Platforms.Any(p => p.Equals(displayName, StringComparison.OrdinalIgnoreCase)))
        {
            member.Platforms.Add(displayName);
        }
    }

    /// <summary>
    /// Opportunistically updates a formal clan member's last-seen game/server from a presence
    /// sighting (joining/showing up in a channel) rather than actual chat — a no-op for anyone
    /// not already tracked, since presence alone shouldn't auto-create a roster entry the way
    /// actually talking does (see RecordSeen), AND a no-op for a tracked-but-informal entry
    /// (IsClanMember false — someone who's merely spoken before, not a real clan member): every
    /// Join/ShowUser/UserFlags event for every tracked user in a channel called this
    /// unconditionally, so a mass-join burst including even a few informal entries (e.g. a load
    /// test's bots, auto-tracked from an earlier run's chat) triggered a full roster disk-write
    /// per event — confirmed live as a real cause of the app hanging under one. Actual clan
    /// members are the only ones this app has ever claimed to track presence for; everyone else
    /// only gets updated from RecordSeen (an actual chat message), which is inherently far rarer
    /// than "reconnected to a channel."
    /// </summary>
    public static void RecordProductSeen(string username, string product, string server)
    {
        lock (SyncRoot)
        {
            var member = Find(username);
            if (member is null || !member.IsClanMember)
            {
                return;
            }

            member.LastSeenProduct = product;
            RecordPlatform(member, product);
            member.LastSeenServer = server;
            MarkDirtyAndNotify();
        }
    }

    /// <summary>Locked so concurrent Save() calls from multiple bots (or a bot and the Clan Members window) can't collide writing the same file. Immediate, unlike the auto-tracking paths (RecordSeen/RecordProductSeen) — this is only ever called for an explicit user action (e.g. the Clan Members window's own Save), where "wrote to disk right now" is the expected behavior. Also the one place an in-place rename (Name/Aliases edited on an existing member, so Members.Count doesn't change) reaches this store, so it explicitly invalidates NameIndex rather than relying on the count-mismatch check.</summary>
    public static void Save()
    {
        lock (SyncRoot)
        {
            SaveLocked();
            _saveDirty = false;
            _nameIndex = null;
        }

        RosterChanged?.Invoke();
    }

    /// <summary>
    /// For a caller that mutated a ClanMember obtained via Find/FindTrusted directly (e.g.
    /// BotEngine's auto-whisper bookkeeping stamping LastAutoWhisperUtc) rather than through one
    /// of this store's own Record* methods — same debounced-write behavior as those, just a
    /// public entry point for external mutation instead of a private implementation detail.
    /// </summary>
    public static void MarkDirty()
    {
        lock (SyncRoot)
        {
            MarkDirtyAndNotify();
        }
    }

    /// <summary>
    /// Marks the roster dirty and schedules (or reschedules) a single debounced disk write
    /// SaveDebounceInterval from now, coalescing a burst of many rapid-fire RecordSeen/
    /// RecordProductSeen calls (e.g. several tracked users joining/talking within the same
    /// couple of seconds) into one JSON-serialize-and-write instead of one per call — the actual
    /// fix for the mass-join hang (see RecordProductSeen's remarks). RosterChanged still fires
    /// immediately so any open UI reflects the in-memory change right away; only the disk I/O is
    /// deferred. Callers already inside `lock (SyncRoot)` (RecordSeen/RecordProductSeen) call this
    /// directly; external callers go through the public MarkDirty() wrapper above instead.
    /// </summary>
    private static void MarkDirtyAndNotify()
    {
        _saveDirty = true;
        _saveTimer ??= new Timer(static _ => FlushIfDirty(), null, Timeout.Infinite, Timeout.Infinite);
        _saveTimer.Change(SaveDebounceInterval, Timeout.InfiniteTimeSpan);
        RosterChanged?.Invoke();
    }

    private static void FlushIfDirty()
    {
        lock (SyncRoot)
        {
            if (_saveDirty)
            {
                SaveLocked();
                _saveDirty = false;
            }
        }
    }

    /// <summary>Forces any pending debounced save out immediately — called on app shutdown so the last couple of seconds of auto-tracked activity aren't silently lost.</summary>
    public static void FlushPendingSave() => FlushIfDirty();

    private static void SaveLocked()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        File.WriteAllText(FilePath, JsonSerializer.Serialize(Members, JsonOptions));
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
