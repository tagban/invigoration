namespace Invigoration.Core.Sc2;

/// <summary>
/// Battle.net program codes as friends lists report them ("S1", "Fen", "BSAp"...), turned into
/// what the Friends tab shows: an icon (a classic product code or an icon key) and a name.
/// </summary>
public static class BattlenetPrograms
{
    public static (string Icon, string Name) Describe(string program) => program switch
    {
        "S1" => ("PXES", "StarCraft: Remastered"),
        "S2" => ("sc2", "StarCraft II"),
        "W3" => ("3RAW", "Warcraft III: Reforged"),
        "W1R" => ("war1", "Warcraft: Remastered"),
        "OSI" => ("d2r", "Diablo II: Resurrected"),
        "D3" => ("d3", "Diablo III"),
        "Fen" => ("d4", "Diablo IV"),
        "ANBS" => ("diabloimmortal", "Diablo Immortal"),
        "WoW" => ("wow", "World of Warcraft"),
        "Pro" => ("overwatch", "Overwatch"),
        "WTCG" => ("hearthstone", "Hearthstone"),
        "Hero" => ("hots", "Heroes of the Storm"),
        "GRY" => ("wcrumble", "Warcraft Rumble"),
        "BSAp" => ("bnet2", "Battle.net mobile app"),
        "App" => ("bnet2", "Battle.net app"),
        "" => ("bnet2", "Online"),
        var other => ("bnet2", other),
    };
}
