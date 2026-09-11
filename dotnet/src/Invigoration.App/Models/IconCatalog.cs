namespace Invigoration.App.Models;

/// <summary>
/// Single source of truth for every editable icon key this app knows about, keyed by category —
/// shared by IconManagerViewModel (the "Manage Icons" editor) and anywhere else that needs to
/// offer a pick-an-icon list (e.g. ConfigViewModel's tab-group icon picker), so the two never
/// drift out of sync.
/// </summary>
public static class IconCatalog
{
    public static readonly (string Key, string DisplayName)[] GameIcons =
    [
        ("sc", "StarCraft"),
        ("scbw", "StarCraft: Brood War"),
        ("jsc", "StarCraft (Japanese release)"),
        ("sware", "StarCraft (Shareware)"),
        ("sc-stars0", "StarCraft Ladder (0 wins / ranked plate)"),
        ("sc-stars1", "StarCraft Ladder (1 win)"),
        ("sc-stars2", "StarCraft Ladder (2 wins)"),
        ("sc-stars3", "StarCraft Ladder (3 wins)"),
        ("sc-stars4", "StarCraft Ladder (4 wins)"),
        ("sc-stars5", "StarCraft Ladder (5 wins)"),
        ("sc-stars6", "StarCraft Ladder (6 wins)"),
        ("sc-stars7", "StarCraft Ladder (7 wins)"),
        ("sc-stars8", "StarCraft Ladder (8 wins)"),
        ("sc-stars9", "StarCraft Ladder (9 wins)"),
        ("sc-stars10", "StarCraft Ladder (10+ wins)"),
        ("war2", "Warcraft II: Battle.net Edition"),
        ("war2-axes0", "Warcraft II Ladder (0 wins)"),
        ("war2-axes1", "Warcraft II Ladder (1 win)"),
        ("war2-axes2", "Warcraft II Ladder (2 wins)"),
        ("war2-axes3", "Warcraft II Ladder (3 wins)"),
        ("war2-axes4", "Warcraft II Ladder (4 wins)"),
        ("war2-axes5", "Warcraft II Ladder (5 wins)"),
        ("war2-axes6", "Warcraft II Ladder (6 wins)"),
        ("war2-axes7", "Warcraft II Ladder (7 wins)"),
        ("war2-axes8", "Warcraft II Ladder (8 wins)"),
        ("war2-sword1", "Warcraft II Ladder (9 wins)"),
        ("war2-sword2", "Warcraft II Ladder (10+ wins)"),
        ("war2-ranked", "Warcraft II Ladder (ranked plate)"),
        ("war3", "Warcraft III"),
        ("w3tft", "Warcraft III: The Frozen Throne"),
        ("diablo", "Diablo"),
        ("diablo-dot0", "Diablo (No kills)"),
        ("diablo-dot1", "Diablo (Normal cleared)"),
        ("diablo-dot2", "Diablo (Nightmare cleared)"),
        ("diablo-dot3", "Diablo (Hell cleared)"),
        ("dshr", "Diablo: Shareware"),
        ("diablo2", "Diablo II"),
        ("d2exp", "Diablo II: Lord of Destruction"),
        ("chat", "Chat Client (generic)"),
        ("sc2", "StarCraft II"),
    ];

    /// <summary>
    /// Modern Battle.net account games with no classic-era chat icon lineage at all (StarCraft II
    /// is the exception — already a GameIcons entry, since it replaces a real classic-style
    /// default) — mostly not connectable products yet, kept ready for whenever a Stimpak-backed
    /// friend/roster entry can report which of these it's actually playing (see BotEngine.Sc2.cs's
    /// HandleSc2FriendsReceived, currently hardcoded to "sc2" for every contact — Stimpak's Friend
    /// data has no per-contact game field to read yet). Assets sourced from the official
    /// account.battle.net game-icon SVGs, rasterized to PNG since nothing in this app renders SVG.
    /// </summary>
    public static readonly (string Key, string DisplayName)[] Bnet2Icons =
    [
        ("wow", "World of Warcraft"),
        ("war1", "Warcraft: Remastered"),
        ("d3", "Diablo III"),
        ("d4", "Diablo IV"),
        ("d2r", "Diablo II: Resurrected"),
        ("diabloimmortal", "Diablo Immortal"),
        ("overwatch", "Overwatch"),
        ("hearthstone", "Hearthstone"),
        ("hots", "Heroes of the Storm"),
        ("wcrumble", "Warcraft Rumble"),
    ];

    public static readonly (string Key, string DisplayName)[] StatusIcons =
    [
        ("blizz", "Blizzard Representative"),
        ("sysop", "Administrator"),
        ("mod-gavel", "Channel Operator"),
        ("mega", "Speaker"),
        ("guest", "Special Guest (VIP)"),
        ("ignore", "Squelched"),
    ];

    public static readonly (string Key, string DisplayName)[] FriendIcons =
    [
        ("offline", "Offline Friend Indicator"),
    ];

    /// <summary>Not a Blizzard product/roster badge at all — small text/emoji badges for things this app itself distinguishes, with no real game icon to draw from.</summary>
    public static readonly (string Key, string DisplayName)[] CustomIcons =
    [
        ("bnet", "Battle.net (classic)"),
        ("bnet2", "Battle.net 2.0"),
        ("pvpgn", "PVPGN"),
        ("atlas", "Atlas"),
        ("test", "Test"),
        ("whisper", "Whispers Tab"),
        ("youtube-music", "YouTube Music"),
        ("hotline", "Hotline"),
        ("discord-relay", "Discord (relay)"),
    ];
}
