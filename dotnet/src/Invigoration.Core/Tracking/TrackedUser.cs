namespace Invigoration.Core.Tracking;

/// <summary>
/// One user seen on a non-Battle.net protocol (Hotline today; IRC/FFXI are the explicitly named
/// future ones) — last-seen and trivia score, kept separate from Clan.ClanRosterStore/ClanMember
/// on purpose per explicit request: these people were never added to a clan roster (there's no
/// "clanadd" equivalent here), so they must never show up in the Battle.net "CLAN" list/commands.
/// Tagging by protocol+server (not just name) matters for the same reason BnetUsername qualifies a
/// name with "@server" — the same nickname on two different Hotline servers (or a future IRC
/// network) is not the same person.
/// </summary>
public sealed class TrackedUser
{
    public string Name { get; set; } = "";

    /// <summary>"Hotline" today; "IRC"/"FFXI" are the named future ones — a plain string, not an enum, so a new protocol never needs a Core-wide schema change to start tracking.</summary>
    public string Protocol { get; set; } = "";

    /// <summary>The specific server/network host this user was seen on (e.g. a Hotline server's host:port) — part of this record's identity, not just descriptive metadata; see Matches.</summary>
    public string Server { get; set; } = "";

    public DateTime LastSeenUtc { get; set; }

    public double TriviaScore { get; set; }

    /// <summary>"Name [Protocol:Server]" — the "add the protocol/server after their name" display form used in leaderboards and lookups, e.g. "Tagban [Hotline:bigredh.com]".</summary>
    public string QualifiedName => $"{Name} [{Protocol}:{Server}]";

    public bool Matches(string name, string protocol, string server) =>
        Name.Equals(name, StringComparison.OrdinalIgnoreCase) &&
        Protocol.Equals(protocol, StringComparison.OrdinalIgnoreCase) &&
        Server.Equals(server, StringComparison.OrdinalIgnoreCase);
}
