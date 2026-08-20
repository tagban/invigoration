namespace Invigoration.Core.Commands;

/// <summary>One command a permission level can be granted access to, with its recognized text aliases.</summary>
public sealed record CommandCatalogEntry(string CanonicalName, string DisplayName, IReadOnlyList<string> Aliases);

/// <summary>
/// The full set of bot commands, grouped by canonical name so aliases (e.g.
/// "h"/"hex", "disc"/"disconnect") share one permission-level entry instead
/// of needing to be granted separately.
/// </summary>
public static class CommandCatalog
{
    public static readonly IReadOnlyList<CommandCatalogEntry> Entries =
    [
        new("idle", "Set idle message", ["idle"]),
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
    ];

    private static readonly Dictionary<string, string> AliasToCanonical =
        Entries.SelectMany(e => e.Aliases.Select(a => (Alias: a, e.CanonicalName)))
            .ToDictionary(x => x.Alias, x => x.CanonicalName, StringComparer.OrdinalIgnoreCase);

    /// <summary>Resolves a typed command word (any alias) to its canonical name, or returns it unchanged if unrecognized.</summary>
    public static string ResolveCanonicalName(string typedCommand) =>
        AliasToCanonical.GetValueOrDefault(typedCommand, typedCommand);
}
