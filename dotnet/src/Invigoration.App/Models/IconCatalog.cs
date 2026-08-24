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
        ("war2", "Warcraft II: Battle.net Edition"),
        ("war3", "Warcraft III"),
        ("w3tft", "Warcraft III: The Frozen Throne"),
        ("diablo", "Diablo"),
        ("dshr", "Diablo: Shareware"),
        ("diablo2", "Diablo II"),
        ("d2exp", "Diablo II: Lord of Destruction"),
        ("chat", "Chat Client (generic)"),
    ];

    public static readonly (string Key, string DisplayName)[] StatusIcons =
    [
        ("blizz", "Blizzard Representative"),
        ("sysop", "Administrator"),
        ("mod-gavel", "Channel Operator"),
        ("mega", "Speaker / VIP"),
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
    ];
}
