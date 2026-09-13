namespace Invigoration.Core.StatString;

/// <summary>
/// A Diablo II character as carried in a VD2D/PX2D chat statstring — the structured form of what
/// <see cref="StatStringParser.ParseD2Stats"/> used to read and immediately flatten into a
/// sentence. Exposed so the channel list can pick the character's real Battle.net portrait tile
/// (see Chat.D2PortraitTile) instead of just describing it.
/// </summary>
/// <remarks>
/// Wire format per bnetdocs.org/document/18/chat-statstrings: <c>ProductID + Realm + ',' +
/// CharacterName + ',' + 33 bytes</c>, or a bare ProductID for an Open character. Byte offsets
/// into the 33-byte struct: 13 class (1-7), 25 level, 26 flags (0x04 hardcore, 0x08 dead, 0x20
/// expansion), 27 current act, 30 ladder (0xFF = non-ladder). The struct arrives Latin-1 decoded
/// (PacketReader), so each char is exactly one wire byte.
/// </remarks>
public readonly record struct D2Character(
    string Realm,
    string Name,
    int ClassIndex,
    int Level,
    bool Hardcore,
    bool Dead,
    bool Expansion,
    bool Ladder,
    int Progress)
{
    public const int ClassicClassCount = 5;
    public const int ExpansionClassCount = 7;

    private const int ClassOffset = 13;
    private const int LevelOffset = 25;
    private const int FlagsOffset = 26;
    private const int ActOffset = 27;
    private const int LadderOffset = 30;

    /// <summary>
    /// Act progression as a 0-based count from the "Current Act" byte: <c>0x80 + 2 × progress</c>.
    /// Classic runs 0-12 (four acts per difficulty, 12 = all acts completed); the expansion runs
    /// 0-15 in steps of five per difficulty (Normal 0-3, Nightmare 5-8, Hell 10-13, 15 = all
    /// completed — bnetdocs lists Act IV and V as sharing one value, so 4/9/14 aren't sent).
    /// </summary>
    public static int ProgressFromActByte(byte actByte) => (actByte & 0x3E) >> 1;

    /// <summary>Which difficulty the character has fully completed, 0-3 — the rank their title reflects.</summary>
    public int TitleTier => Math.Clamp(Progress / (Expansion ? 5 : 4), 0, 3);

    /// <summary>Whether this character's product is the expansion (PX2D) — decides the label, independent of <see cref="Expansion"/>, which is the character's own flag.</summary>
    public static bool IsD2Product(string statString) =>
        statString.StartsWith("VD2D", StringComparison.Ordinal) || statString.StartsWith("PX2D", StringComparison.Ordinal);

    /// <summary>Parses a D2 statstring's character struct. False for any other product, an Open character, or a struct too short to hold the class/level/flags/act bytes.</summary>
    public static bool TryParse(string statString, out D2Character character)
    {
        character = default;
        if (!IsD2Product(statString))
        {
            return false;
        }

        var firstComma = statString.IndexOf(',', 4);
        if (firstComma < 0)
        {
            return false;
        }

        var secondComma = statString.IndexOf(',', firstComma + 1);
        if (secondComma < 0)
        {
            return false;
        }

        var p = statString.AsSpan(secondComma + 1);
        if (p.Length <= ActOffset)
        {
            return false;
        }

        var flags = (byte)p[FlagsOffset];
        character = new D2Character(
            Realm: statString[4..firstComma],
            Name: statString[(firstComma + 1)..secondComma],
            ClassIndex: (byte)p[ClassOffset] - 1,
            Level: (byte)p[LevelOffset],
            Hardcore: (flags & 0x04) != 0,
            Dead: (flags & 0x08) != 0,
            // An expansion character can only exist on the expansion product, whatever the byte says.
            Expansion: statString.StartsWith("PX2D", StringComparison.Ordinal) && (flags & 0x20) != 0,
            Ladder: p.Length > LadderOffset && (byte)p[LadderOffset] != 0xFF,
            Progress: ProgressFromActByte((byte)p[ActOffset]));
        return true;
    }
}
