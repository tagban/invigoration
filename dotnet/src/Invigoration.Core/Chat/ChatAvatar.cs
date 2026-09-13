using Invigoration.Core.StatString;

namespace Invigoration.Core.Chat;

/// <summary>
/// Which Battle.net chat avatar Diablo II's lobby draws for a user — the full-body animated figures
/// along the bottom of its chat screen. Real D2 characters appear as themselves; everyone else gets
/// one of the fixed avatars documented on The Arreat Summit
/// (classic.battle.net/diablo2exp/basics/bnetchat.shtml, where the animations themselves come from).
/// Like every other icon, the server only says who someone is (flags + statstring); the client
/// picks the picture.
/// </summary>
public static class ChatAvatar
{
    public const string BlizzardRep = "blizzrep";
    public const string Sysop = "sysop";
    public const string Moderator = "moderator";
    public const string Speaker = "speaker";
    public const string Referee = "referee";
    public const string DeadHardcore = "deadhardcore";
    public const string Unknown = "unknown";
    public const string ChatClient = "chatclient";
    public const string StarCraft = "starcraft";
    public const string BroodWar = "broodwar";
    public const string WarcraftII = "war2";
    public const string Diablo = "diablo";

    /// <summary>
    /// Battle.net's tournament-official flags (PGL official 0x2000, WCG official 0x8000, GF official
    /// 0x100000, per bnetdocs). The Arreat Summit only says the Referee is "authorized for disputes"
    /// without naming a flag, so tying it to these is an inference — they're the flags that marked
    /// tournament staff.
    /// </summary>
    private const uint RefereeFlags = 0x2000 | 0x8000 | 0x100000;

    /// <summary>
    /// The avatar key for this user, or null when they're a live Diablo II character who should be
    /// drawn as themselves. Rank comes first — a D2 player running a channel shows as the Moderator,
    /// with the ban hammer — then the client they're on. Open (non-realm) D2 characters carry no
    /// character data and appear as the Unknown avatar, as they did in the real client.
    /// </summary>
    public static string? For(uint flags, string statString)
    {
        var f = (UserFlags)flags;
        if (f.HasFlag(UserFlags.Blizzard))
        {
            return BlizzardRep;
        }

        if (f.HasFlag(UserFlags.Admin))
        {
            return Sysop;
        }

        if (f.HasFlag(UserFlags.Operator))
        {
            return Moderator;
        }

        if (f.HasFlag(UserFlags.Speaker))
        {
            return Speaker;
        }

        if ((flags & RefereeFlags) != 0)
        {
            return Referee;
        }

        if (D2Character.IsD2Product(statString))
        {
            if (!D2Character.TryParse(statString, out var character))
            {
                return Unknown;
            }

            return character.Hardcore && character.Dead ? DeadHardcore : null;
        }

        var product = statString.Length >= 4 ? statString[..4] : statString;
        return product switch
        {
            "RATS" or "RTSJ" or "RHSS" => StarCraft,
            "PXES" => BroodWar,
            "NB2W" => WarcraftII,
            "LTRD" or "RHSD" => Diablo,
            "TAHC" => ChatClient,
            _ => Unknown,
        };
    }
}
