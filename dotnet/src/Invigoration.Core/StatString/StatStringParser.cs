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
                if (statString.Length > 4)
                {
                    return "";
                }

                return statString.Length == 4
                    ? "WarCraft III: Reign of Chaos (No stats available)"
                    : $"WarCraft III: Reign of Chaos (error: {statString})";

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

    private static string ParseDiabloClassicStats(string statString, string fallback, string label)
    {
        var values = statString.Length > 5 ? statString[5..].Split(' ') : [];
        if (values.Length != 9)
        {
            return fallback;
        }

        var className = values[2] switch
        {
            "0" => "warrior",
            "1" => "rogue",
            "2" => "sorceror",
            _ => "unknown class",
        };

        return $"{label}: (Level {values[0]} {className} with {values[1]} dots, {values[3]} strength, " +
               $"{values[4]} magic, {values[5]} dexterity, {values[6]} vitality, and {values[7]} gold).";
    }

    private static readonly string[] D2Classes =
        ["amazon", "sorceress", "necromancer", "paladin", "barbarian", "druid", "assassin", "unknown class"];

    public static string ParseD2Stats(string stats)
    {
        var header = stats.Length > 4 ? stats[..4] : stats;
        var label = header == "VD2D" ? "Diablo II" : "Diablo II Lord of Destruction";

        if (stats.Length == 4)
        {
            return $"{label}: (Open Character).";
        }

        var firstComma = stats.IndexOf(',', 4);
        if (firstComma < 0)
        {
            return $"{label}: (Open Character).";
        }

        var realm = stats[4..firstComma];
        var secondComma = stats.IndexOf(',', firstComma + 1);
        if (secondComma < 0)
        {
            return $"{label}: (Open Character).";
        }

        var name = stats[(firstComma + 1)..secondComma];
        var p = stats[(secondComma + 1)..];
        if (p.Length <= 27)
        {
            return $"{label}: (Open Character).";
        }

        var charClass = (byte)(p[13] - 1);
        if (charClass > 6)
        {
            charClass = 7;
        }

        var female = charClass is 0 or 1 or 6;
        var charLevel = (byte)p[25];
        var hardcore = ((byte)p[26] & 0x4) != 0;
        var dead = ((byte)p[26] & 0x8) != 0;
        var tier = ((byte)p[27] & 0x18) >> 3;
        var isExpansion = header == "PX2D" && ((byte)p[26] & 0x20) != 0;

        string title = "";
        if (isExpansion)
        {
            title = tier switch
            {
                1 => hardcore ? "Destroyer" : "Slayer",
                2 => hardcore ? "Conquerer" : "Champion",
                3 => hardcore ? "Guardian" : (female ? "Matriarch" : "Patriarch"),
                _ => "",
            };
        }
        else
        {
            title = tier switch
            {
                1 => female ? (hardcore ? "Countess" : "Dame") : (hardcore ? "Count" : "Sir"),
                2 => female ? (hardcore ? "Duchess" : "Lady") : (hardcore ? "Duke" : "Lord"),
                3 => female ? (hardcore ? "Queen" : "Baroness") : (hardcore ? "King" : "Baron"),
                _ => "",
            };
        }

        var titlePrefix = title.Length > 0 ? title + " " : "";
        var deadPrefix = hardcore && dead ? "dead " : "";
        var levelWord = hardcore ? "hardcore level" : "level";

        return $"{label}: ({titlePrefix}{name} a {deadPrefix}{levelWord} {charLevel} {D2Classes[charClass]} on realm {realm}).";
    }
}
