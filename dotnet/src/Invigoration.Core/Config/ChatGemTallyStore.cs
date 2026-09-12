using System.Text.Json;

namespace Invigoration.Core.Config;

/// <summary>One account's chat-gem tally for one calendar month.</summary>
public sealed class ChatGemMonthTally
{
    /// <summary>"yyyy-MM" the count belongs to. The tally resets by starting a new month's entry, not by zeroing this one, so a late submission can still name the month it came from.</summary>
    public string Month { get; set; } = "";

    /// <summary>Activations counted locally this month.</summary>
    public int Count { get; set; }

    /// <summary>The highest Count the server has acknowledged. Submissions send the cumulative Count, never a delta — a lost send then costs nothing, since the next one carries the full figure and the server can take the max. See ChatGemTallySender.</summary>
    public int Submitted { get; set; }
}

/// <summary>
/// The opt-in chat-gem leaderboard's local state: whether this install agreed to share at all,
/// and the per-account monthly tallies waiting to be submitted. Persisted at
/// %AppData%/Invigoration/chat-gem-tally.json, following the same static-store shape as
/// TabGroupIconStore.
///
/// Deliberately separate from BotConfig.ChatGemActivations, which stays an all-time per-bot
/// counter for the gem's own tooltip: this store is the sharing side of the feature and resets
/// monthly, so conflating the two would make the tooltip lie every 1st of the month.
///
/// Nothing here transmits anything — see ChatGemTallySender, which is inert until an endpoint is
/// configured.
/// </summary>
public static class ChatGemTallyStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private static readonly Lock SyncRoot = new();
    private static ChatGemTallyState? _cache;
    private static string? _configDirectoryOverride;

    /// <summary>Test-only hook, same pattern as TabGroupIconStore's — redirects reads/writes to an isolated directory instead of the real %AppData%/Invigoration.</summary>
    public static string? ConfigDirectoryOverride
    {
        get => _configDirectoryOverride;
        set
        {
            _configDirectoryOverride = value;
            _cache = null;
        }
    }

    private static string FilePath => Path.Combine(_configDirectoryOverride ?? ConfigStore.DefaultConfigDirectory(), "chat-gem-tally.json");

    private static ChatGemTallyState State => _cache ??= LoadFromDisk();

    /// <summary>Null until the user has been asked (the prompt fires on their first ever gem click); true/false once they've answered. A "no" is remembered so they're never asked twice.</summary>
    public static bool? ShareConsent
    {
        get => State.ShareConsent;
        set
        {
            State.ShareConsent = value;
            Save();
        }
    }

    /// <summary>True only once the user has actively agreed — the default (unanswered) shares nothing.</summary>
    public static bool SharingEnabled => State.ShareConsent == true;

    /// <summary>
    /// Counts one activation for an account, rolling the tally over if the month has changed
    /// since its last activation. <paramref name="accountKey"/> is the "username@server" identity
    /// the leaderboard is keyed by. Counting happens regardless of consent — the count is local
    /// either way, and only submission is gated — so that answering "yes" later doesn't start
    /// them from zero mid-month.
    /// </summary>
    public static void RecordActivation(string accountKey, DateTimeOffset now)
    {
        if (string.IsNullOrWhiteSpace(accountKey))
        {
            return;
        }

        var month = MonthKey(now);
        if (!State.Tallies.TryGetValue(accountKey, out var tally) || tally.Month != month)
        {
            tally = new ChatGemMonthTally { Month = month };
            State.Tallies[accountKey] = tally;
        }

        tally.Count++;
        Save();
    }

    /// <summary>Every account whose current count is ahead of what the server has acknowledged. Empty when there's nothing new to send.</summary>
    public static IReadOnlyList<(string AccountKey, ChatGemMonthTally Tally)> PendingSubmissions() =>
        [.. State.Tallies.Where(kv => kv.Value.Count > kv.Value.Submitted).Select(kv => (kv.Key, kv.Value))];

    /// <summary>Records that the server acknowledged this cumulative figure. Never moves Submitted backwards, so an out-of-order acknowledgement can't cause a resend loop.</summary>
    public static void MarkSubmitted(string accountKey, int count)
    {
        if (State.Tallies.TryGetValue(accountKey, out var tally) && count > tally.Submitted)
        {
            tally.Submitted = count;
            Save();
        }
    }

    /// <summary>This month's local count for an account, or 0 if it has none yet.</summary>
    public static int CurrentMonthCount(string accountKey, DateTimeOffset now) =>
        State.Tallies.TryGetValue(accountKey, out var tally) && tally.Month == MonthKey(now) ? tally.Count : 0;

    private static string MonthKey(DateTimeOffset now) => now.ToString("yyyy-MM");

    private static void Save()
    {
        lock (SyncRoot)
        {
            var directory = Path.GetDirectoryName(FilePath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(FilePath, JsonSerializer.Serialize(State, JsonOptions));
        }
    }

    private static ChatGemTallyState LoadFromDisk()
    {
        try
        {
            return File.Exists(FilePath)
                ? JsonSerializer.Deserialize<ChatGemTallyState>(File.ReadAllText(FilePath)) ?? new ChatGemTallyState()
                : new ChatGemTallyState();
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            // A corrupt or unreadable tally file is not worth failing a chat client over — start
            // fresh rather than throwing on a purely cosmetic leaderboard.
            return new ChatGemTallyState();
        }
    }

    /// <summary>Test-only: drops the in-memory cache so the next read comes from disk.</summary>
    public static void ResetCacheForTests() => _cache = null;
}

/// <summary>The on-disk shape of <see cref="ChatGemTallyStore"/>.</summary>
public sealed class ChatGemTallyState
{
    public bool? ShareConsent { get; set; }

    public Dictionary<string, ChatGemMonthTally> Tallies { get; set; } = [];
}
