using Invigoration.Core.Chat;

namespace Invigoration.Core.Clan;

/// <summary>
/// One tracked person in the shared clan roster. Rank is a free-form label
/// (e.g. "Officer", "Recruit") the bot master defines themselves — it isn't
/// tied to Battle.net's own channel flags (Operator/Speaker/etc.), since the
/// point is an org structure that survives someone changing which Battle.net
/// account they're logged into.
/// </summary>
public sealed class ClanMember
{
    /// <summary>Primary/display Battle.net username — an actual account name, used for matching (see Matches). For a personal label instead, use NickName.</summary>
    public string Name { get; set; } = "";

    /// <summary>Freeform personal label (e.g. "John") for display and search — never used to match a speaking user. Name/Aliases (real Battle.net account names) are what matching runs against.</summary>
    public string NickName { get; set; } = "";

    public string Rank { get; set; } = "";

    /// <summary>
    /// True for members explicitly added via "Add Member" or the "clanadd"
    /// command — false for entries auto-created just from someone talking
    /// while ClanRosterStore.RecordSeen's default-rank auto-registration is
    /// on. Lets the UI filter the roster down to "just clan members" versus
    /// the full ongoing seen-list of everyone the bot has observed.
    /// </summary>
    public bool IsClanMember { get; set; }

    /// <summary>Wire product code (e.g. "PX2D") of the last game this member was seen playing, if known — drives which icon the UI shows for them.</summary>
    public string LastSeenProduct { get; set; } = "";

    /// <summary>Battle.net server host (e.g. "useast.battle.net") this member was last seen on.</summary>
    public string LastSeenServer { get; set; } = "";

    /// <summary>Other Battle.net usernames this same person might be seen on.</summary>
    public List<string> Aliases { get; set; } = [];

    public string Notes { get; set; } = "";

    /// <summary>UTC timestamp of the last time the bot saw this member (by primary name or alias) talk, emote, or whisper, across any bot tab. Null if never observed.</summary>
    public DateTime? LastSeenUtc { get; set; }

    /// <summary>UTC timestamp of the last time an auto-whisper (per their rank's ClanRank.AutoWhisperFrequency) was sent to this member — used to enforce Daily/Once frequency limits. Null if never sent.</summary>
    public DateTime? LastAutoWhisperUtc { get; set; }

    /// <summary>Running trivia-game score, adjustable via the "score" command; 0 until trivia is played (or scores are set manually). Fractional since correct answers award graduated points based on how many hints were shown (see BotConfig.TriviaPointsBeforeFirstHint and friends).</summary>
    public double TriviaScore { get; set; }

    /// <summary>
    /// Optional free-form platform labels (e.g. "Classic", "SC2", "SC:R")
    /// this member is known to play on — purely organizational. Matching and
    /// last-seen tracking are already platform-agnostic (any bot, on any
    /// product, resolves against this one shared roster by username), so
    /// this doesn't gate anything; it just lets a clan spanning multiple
    /// Battle.net generations note where each member is actually seen.
    /// </summary>
    public List<string> Platforms { get; set; } = [];

    /// <summary>
    /// True if the given Battle.net username is this member's primary name
    /// or one of their aliases — normalized via <see cref="BnetUsername"/>
    /// so a Diablo II player showing as "*Name" (in-game) still matches.
    /// Unscoped by server: any "name@server" qualifier on the stored side is
    /// stripped before comparing, so this matches regardless of which server
    /// the entry happens to be pinned to — for editing/management lookups
    /// where the operator is naming an account directly, not for
    /// authorization decisions (use <see cref="MatchesOnServer"/> for those,
    /// which is where the qualifier actually gets enforced).
    /// </summary>
    public bool Matches(string username) =>
        BnetUsername.Equals(BnetUsername.SplitServerQualifier(Name).Name, username) ||
        Aliases.Any(a => BnetUsername.Equals(BnetUsername.SplitServerQualifier(a).Name, username));

    /// <summary>
    /// Server-scoped version of <see cref="Matches"/> for authorization
    /// decisions (bot-master check, rank-based permission grants, ban
    /// checks): a Name/Alias qualified as "name@server" only matches a
    /// speaker actually on that server; an unqualified one matches any
    /// server, same as <see cref="Matches"/>. See
    /// <see cref="BnetUsername.MatchesOnServer"/> for why this exists —
    /// classic Battle.net accounts are per-gateway, so the same bare name
    /// can be a completely different, unrelated person on another server.
    /// </summary>
    public bool MatchesOnServer(string username, string speakerServer) =>
        BnetUsername.MatchesOnServer(username, Name, speakerServer) ||
        Aliases.Any(a => BnetUsername.MatchesOnServer(username, a, speakerServer));
}
