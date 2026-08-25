namespace Invigoration.Core.Commands;

/// <summary>
/// One command a permission level can be granted access to, with its recognized text aliases.
/// Usage (optional — most commands are self-explanatory from DisplayName alone) is shown when
/// the command's own handler replies with it, typically because it was typed bare/with missing
/// args — see BotEngine.Commands.cs's ReplyWithUsageAsync and HandleHelpCommandAsync, which also
/// surfaces it for any command via "help &lt;command&gt;"/"? &lt;command&gt;".
/// </summary>
public sealed record CommandCatalogEntry(string CanonicalName, string DisplayName, IReadOnlyList<string> Aliases, string? Usage = null);

/// <summary>
/// The full set of bot commands, grouped by canonical name so aliases (e.g.
/// "h"/"hex", "disc"/"disconnect") share one permission-level entry instead
/// of needing to be granted separately.
/// </summary>
public static class CommandCatalog
{
    public static readonly IReadOnlyList<CommandCatalogEntry> Entries =
    [
        new("help", "Show command usage", ["help", "?"],
            "help <command> — shows how to use <command> (works with any of its aliases too). \"help\" alone shows this."),
        new("idle", "Set idle message", ["idle"],
            "idle <minutes> <message> — sends <message> once after <minutes> of no chat activity. " +
            "\"idle off\" turns it off. Placeholders: %Ver%, %Uptime%, %MusicPlaying%, %Username%."),
        new("disconnect", "Disconnect", ["disconnect", "disc"]),
        new("colors", "Show color-code help", ["colors", "color"]),
        new("reconnect", "Reconnect", ["reconnect"]),
        new("hex", "Send hex-obfuscated text", ["hex", "h"]),
        new("invigencrypt", "Send Invig-encrypted text", ["invigencrypt", "encrypt", "ie", "i"]),
        new("sysinfo", "Show system info", ["sysinfo"]),
        new("ver", "Show bot version", ["ver"]),
        new("uptime", "Show uptime", ["uptime"]),
        new("about", "Show about info", ["about"]),
        new("say", "Say raw text", ["say"]),
        new("bancount", "Show ban count", ["bancount"]),
        new("kickcount", "Show kick count", ["kickcount"]),
        new("joincount", "Show join count", ["joincount"]),
        new("ban", "Ban a user", ["ban"]),
        new("kick", "Kick a user", ["kick"]),
        new("join", "Join a channel", ["join"]),
        new("user", "Set whisper-focus user", ["user"]),
        new("useroff", "Clear whisper-focus user", ["useroff"]),
        new("prepend", "Set prepend text", ["prepend", "pre"]),
        new("postpend", "Set postpend text", ["postpend", "post"]),
        new("setmaster", "Change bot master", ["setmaster"]),
        new("sethome", "Change home channel", ["sethome"]),
        new("setusername", "Change login username", ["setusername"]),
        new("setpass", "Change login password", ["setpass"]),
        new("setserver", "Change Battle.net server", ["setserver"]),
        new("settrigger", "Change command trigger", ["settrigger"]),
        new("trigger", "Show current trigger", ["trigger"]),
        new("lastreceived", "Show last whisper received", ["last", "lastm", "lastw", "lrm", "lrw"]),
        new("lastsent", "Show last whisper sent", ["lastsend", "lastsm", "lastsw", "lsm", "lsw"]),
        new("canada", "Toggle Canada mode", ["canada"]),
        new("accept", "Toggle clan-invite auto-accept", ["accept"]),
        new("debug", "Toggle debug logging", ["debug"]),
        new("leetspeak", "Toggle leetspeak mode", ["leetspeak"]),
        new("fudd", "Toggle Elmer Fudd mode", ["fudd"]),
        new("moo", "Toggle moo mode", ["moo"]),
        new("home", "Rejoin home channel", ["home", "gohome", "homechan", "homechannel"]),
        new("clanadd", "Add/update a clan member", ["clanadd"]),
        new("clanremove", "Remove a clan member", ["clanremove"]),
        new("clanrank", "Change a clan member's rank", ["clanrank"]),
        new("clanalias", "Add a clan member alias", ["clanalias"]),
        new("clanunalias", "Remove a clan member alias", ["clanunalias"]),
        new("claninfo", "Show a clan member's info", ["claninfo"]),
        new("clanlist", "List clan members", ["clanlist"]),
        new("clanscore", "Adjust a clan member's trivia score", ["clanscore"]),
        new("trivia", "Start/stop the trivia game", ["trivia"]),
        new("musicskip", "Skip to the next track", ["skip", "next"],
            "skip (or next) — skips to the next track on whichever music service is open. No arguments."),
        new("musicthumbsup", "Like the current track", ["thumbsup"],
            "thumbsup — likes the current track. No arguments. Quietly does nothing on a service with no \"like\" concept."),
        new("musicthumbsdown", "Dislike the current track", ["thumbsdown"],
            "thumbsdown — dislikes the current track. No arguments. Quietly does nothing on a service with no \"dislike\" concept (e.g. Spotify)."),
        new("nowplaying", "Show the current track", ["nowplaying", "np", "music"],
            "nowplaying (or np, music) — replies with the currently-playing track and service. No arguments."),
    ];

    private static readonly Dictionary<string, string> AliasToCanonical =
        Entries.SelectMany(e => e.Aliases.Select(a => (Alias: a, e.CanonicalName)))
            .ToDictionary(x => x.Alias, x => x.CanonicalName, StringComparer.OrdinalIgnoreCase);

    private static readonly Dictionary<string, CommandCatalogEntry> ByAlias =
        Entries.SelectMany(e => e.Aliases.Select(a => (Alias: a, Entry: e)))
            .ToDictionary(x => x.Alias, x => x.Entry, StringComparer.OrdinalIgnoreCase);

    /// <summary>Resolves a typed command word (any alias) to its canonical name, or returns it unchanged if unrecognized.</summary>
    public static string ResolveCanonicalName(string typedCommand) =>
        AliasToCanonical.GetValueOrDefault(typedCommand, typedCommand);

    /// <summary>Looks up a command (by canonical name or any alias) and returns its Usage text, or null if unrecognized or it has none defined yet.</summary>
    public static string? GetUsage(string typedCommand) =>
        ByAlias.TryGetValue(typedCommand, out var entry) ? entry.Usage : null;

    /// <summary>
    /// The closest known alias to an unrecognized typed command, for a "did you mean ...?" hint —
    /// null if nothing is close enough to be a plausible typo rather than just an unrelated word
    /// (a real raw server command like "whois"/"f"/"join" typed locally shouldn't get "corrected"
    /// into some unrelated bot command). Distance threshold scales with word length: 1 for very
    /// short words (2-3 chars, where almost anything is "close" and false positives are likely),
    /// otherwise up to 2 — generous enough to catch real typos (transposed/missing/extra letter)
    /// without matching words that just happen to share a couple of letters.
    /// </summary>
    public static string? SuggestClosestAlias(string typedCommand)
    {
        if (typedCommand.Length < 2)
        {
            return null;
        }

        var maxDistance = typedCommand.Length <= 3 ? 1 : 2;
        string? best = null;
        var bestDistance = int.MaxValue;

        foreach (var alias in ByAlias.Keys)
        {
            var distance = LevenshteinDistance(typedCommand, alias);
            if (distance < bestDistance)
            {
                bestDistance = distance;
                best = alias;
            }
        }

        return best is not null && bestDistance <= maxDistance && bestDistance > 0 ? best : null;
    }

    private static int LevenshteinDistance(string a, string b)
    {
        a = a.ToLowerInvariant();
        b = b.ToLowerInvariant();
        var previous = new int[b.Length + 1];
        var current = new int[b.Length + 1];

        for (var j = 0; j <= b.Length; j++)
        {
            previous[j] = j;
        }

        for (var i = 1; i <= a.Length; i++)
        {
            current[0] = i;
            for (var j = 1; j <= b.Length; j++)
            {
                var cost = a[i - 1] == b[j - 1] ? 0 : 1;
                current[j] = Math.Min(Math.Min(current[j - 1] + 1, previous[j] + 1), previous[j - 1] + cost);
            }

            (previous, current) = (current, previous);
        }

        return previous[b.Length];
    }
}
