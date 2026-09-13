namespace Invigoration.Core.StatString;

/// <summary>
/// Turns a user's raw BNCS statstring (from a chat event's Ping/statstring
/// field) into a human-readable description, e.g. "Diablo II: (Level 42
/// sorceress on realm USEast)". Port of modstatstring.bas's ParseStatString
/// and ParseD2Stats.
/// </summary>
public static class StatStringParser
{
    public static string Parse(string statString)
    {
        if (statString.Length < 4)
        {
            return "";
        }

        var product = statString[..4];
        switch (product)
        {
            case "3RAW":
                return ParseWar3Stats(statString, "WarCraft III: Reign of Chaos");

            case "PX3W":
                return ParseWar3Stats(statString, "WarCraft III: The Frozen Throne");

            case "RHSS":
                return "Starcraft Shareware.";

            case "RATS":
                return ParseIconClassStats(statString, "Starcraft");

            case "PXES":
                return ParseIconClassStats(statString, "Starcraft Brood War");

            case "RTSJ":
                return ParseIconClassStats(statString, "Starcraft Japanese");

            case "NB2W":
                return ParseIconClassStats(statString, "Warcraft II");

            case "RHSD":
                return ParseDiabloClassicStats(statString, "a Diablo shareware bot.", "Diablo shareware");

            case "LTRD":
                return ParseDiabloClassicStats(statString, "a Diablo bot.", "Diablo");

            case "PX2D":
            case "VD2D":
                return ParseD2Stats(statString);

            case "TAHC":
                return "a Chat bot.";

            default:
                return "";
        }
    }

    // Contrary to an earlier assumption in this codebase (this case used to unconditionally
    // return "No stats available" for a bare 4-char product code and "" — silently discarding the
    // data — for anything longer), WarCraft III/TFT statstrings do carry real per-user data:
    // bnetdocs.org/document/18/chat-statstrings documents 2 fields plus 1 optional one, space-
    // delimited: an Icon Code ("Level + Tier + \"3W\"", e.g. "5H3W" = level 5 Human), the player's
    // raw level (their highest across all game types; '0' = no ladder games), and an optional
    // reversed clan tag. There can also be 0 fields at all ("often appears with bots who join a
    // channel automatically and not waiting until the user clicks 'Enter Chat'"), and a
    // documented edge case where a statstring carries a level and clan tag but no icon code —
    // handled here by checking whether the first field actually looks like an icon code (ends in
    // "3W") rather than assuming a fixed position. Field-start offset (index 5, same convention
    // as every other product below) was an inferred guess as of 2026-09-11, confirmed correct the
    // next day against the user's own real WC3: TFT connection.
    private static string ParseWar3Stats(string statString, string label)
    {
        var fields = statString.Length > 5 ? statString[5..].Split(' ') : [];
        if (fields.Length == 0 || fields[0].Length == 0)
        {
            return $"{label} (No stats available)";
        }

        var hasIconCode = fields[0].EndsWith("3W", StringComparison.Ordinal);
        var levelField = hasIconCode ? fields.ElementAtOrDefault(1) : fields[0];
        var clanTagField = hasIconCode ? fields.ElementAtOrDefault(2) : fields.ElementAtOrDefault(1);

        var levelSuffix = levelField is null or "0" ? "" : $" (level {levelField})";
        var clanSuffix = clanTagField is { Length: > 0 } ? $" [{new string(clanTagField.Reverse().ToArray())}]" : "";
        return $"{label}{levelSuffix}{clanSuffix}";
    }

    private static string ParseIconClassStats(string statString, string label)
    {
        var values = statString.Length > 5 ? statString[5..].Split(' ') : [];
        if (values.Length != 9)
        {
            var spawnSuffix = values.Length > 3 && values[3] == "1" ? " (spawn)" : "";
            return $"a {label}{spawnSuffix} bot.";
        }

        var spawn = values[3] == "1" ? " (spawn)" : "";
        var wins = values[2];
        var rating = values[0];
        return rating != "0"
            ? $"{label}{spawn}: ({wins} wins, with a rating of {rating} on the ladder)."
            : $"{label}{spawn}: ({wins} wins).";
    }

    // Field order (bnetdocs.org/document/18/chat-statstrings, "Diablo I"): Level, Class, Dots,
    // Strength, Magic, Dexterity, Vitality, Gold, Spawned — confirmed 2026-09-11. This previously
    // read values[1] as dots and values[2] as class (swapped), a real bug: any Diablo player who'd
    // killed Diablo at all showed the wrong class name and a garbled dots value pulled from their
    // actual character class digit instead.
    private static string ParseDiabloClassicStats(string statString, string fallback, string label)
    {
        var values = statString.Length > 5 ? statString[5..].Split(' ') : [];
        if (values.Length != 9)
        {
            return fallback;
        }

        var className = values[1] switch
        {
            "0" => "warrior",
            "1" => "rogue",
            "2" => "sorceror",
            _ => "unknown class",
        };

        return $"{label}: (Level {values[0]} {className} with {values[2]} dots, {values[3]} strength, " +
               $"{values[4]} magic, {values[5]} dexterity, {values[6]} vitality, and {values[7]} gold).";
    }

    private static readonly string[] D2Classes =
        ["amazon", "sorceress", "necromancer", "paladin", "barbarian", "druid", "assassin", "unknown class"];

    public static string ParseD2Stats(string stats)
    {
        var header = stats.Length > 4 ? stats[..4] : stats;
        var label = header == "VD2D" ? "Diablo II" : "Diablo II Lord of Destruction";

        if (!D2Character.TryParse(stats, out var character))
        {
            return $"{label}: (Open Character).";
        }

        var charClass = character.ClassIndex is >= 0 and <= 6 ? character.ClassIndex : 7;
        var female = charClass is 0 or 1 or 6;
        var hardcore = character.Hardcore;

        // TitleTier counts completed difficulties from the act byte. It used to be read straight
        // off bits 0x18 of that byte, which only lines up with the classic game's four acts per
        // difficulty — an expansion character in Nightmare Act IV/V or Hell Act IV/V was already
        // titled for the next difficulty up.
        string title;
        if (character.Expansion)
        {
            title = character.TitleTier switch
            {
                1 => hardcore ? "Destroyer" : "Slayer",
                2 => hardcore ? "Conquerer" : "Champion",
                3 => hardcore ? "Guardian" : (female ? "Matriarch" : "Patriarch"),
                _ => "",
            };
        }
        else
        {
            title = character.TitleTier switch
            {
                1 => female ? (hardcore ? "Countess" : "Dame") : (hardcore ? "Count" : "Sir"),
                2 => female ? (hardcore ? "Duchess" : "Lady") : (hardcore ? "Duke" : "Lord"),
                3 => female ? (hardcore ? "Queen" : "Baroness") : (hardcore ? "King" : "Baron"),
                _ => "",
            };
        }

        var titlePrefix = title.Length > 0 ? title + " " : "";
        var deadPrefix = hardcore && character.Dead ? "dead " : "";
        var levelWord = hardcore ? "hardcore level" : "level";

        return $"{label}: ({titlePrefix}{character.Name} a {deadPrefix}{levelWord} {character.Level} {D2Classes[charClass]} on realm {character.Realm}).";
    }
}
