namespace Invigoration.Core.Chat;

/// <summary>
/// Helpers for comparing Battle.net usernames the way they actually appear
/// on the wire. Diablo II/Lord of Destruction prefixes a player's name with
/// '*' while they're in a game (as opposed to just sitting in the chat
/// channel) — so "PlayerName" and "*PlayerName" are the same account and
/// must compare equal, or setting a bot master/permission-level user/clan
/// alias would silently stop matching the moment that person starts a game.
/// </summary>
public static class BnetUsername
{
    public static string Normalize(string username) => username.StartsWith('*') ? username[1..] : username;

    public static bool Equals(string a, string b) =>
        Normalize(a).Equals(Normalize(b), StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Splits "name@server" into (name, server) — server is null if there's
    /// no '@' (unscoped: matches on any server). Lets a configured identity
    /// (BotMaster, a clan member's name/alias, a permission level's user
    /// list) be pinned to a specific Battle.net server.
    /// </summary>
    public static (string Name, string? Server) SplitServerQualifier(string entry)
    {
        var at = entry.LastIndexOf('@');
        return at > 0 ? (entry[..at], entry[(at + 1)..]) : (entry, null);
    }

    /// <summary>
    /// True if <paramref name="speaker"/> — a username actually seen
    /// chatting on <paramref name="speakerServer"/> — matches
    /// <paramref name="configuredEntry"/>, a BotMaster/alias/permission-user
    /// string that may optionally be "name@server"-qualified.
    ///
    /// Classic Battle.net accounts are scoped per gateway — "tagban" on
    /// useast.battle.net and "tagban" on asia.battle.net are unrelated
    /// accounts that happen to share a name, not the same person. A bare
    /// (unqualified) configured entry matches on any server, same as
    /// before this existed — so existing single-server setups see no
    /// behavior change. Qualifying an entry as "tagban@useast.battle.net"
    /// means it ONLY matches a speaker on that server, so an unrelated
    /// same-named account elsewhere can never be granted whatever that
    /// entry was trusted with (bot-master access, a clan rank, a
    /// permission-level grant).
    /// </summary>
    public static bool MatchesOnServer(string speaker, string configuredEntry, string speakerServer)
    {
        var (name, requiredServer) = SplitServerQualifier(configuredEntry);
        if (!Equals(speaker, name))
        {
            return false;
        }

        return requiredServer is null || ServersMatch(requiredServer, speakerServer);
    }

    private static bool ServersMatch(string a, string b) =>
        NormalizeServer(a).Equals(NormalizeServer(b), StringComparison.OrdinalIgnoreCase);

    /// <summary>Strips a trailing ".battle.net" and folds case, so "useast", "USEast", and "useast.battle.net" are all treated as the same server for qualifier matching.</summary>
    private static string NormalizeServer(string server)
    {
        server = server.Trim();
        const string suffix = ".battle.net";
        if (server.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
        {
            server = server[..^suffix.Length];
        }

        return server.ToLowerInvariant();
    }
}
