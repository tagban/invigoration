using Avalonia.Platform;
using Invigoration.Core.Config;

namespace Invigoration.App.Models;

/// <summary>
/// The icon sets a user can pick: the three bundled with the app, then any they've saved
/// (<see cref="IconSetStore"/>). Applying one copies its images in as the icon overrides, which
/// every icon lookup reads, so one set shows at a time. A bot can name its own
/// (BotConfig.IconSetName), applied whenever its tab is selected.
/// </summary>
public static class IconSets
{
    public const string Bnet1ClassicSetName = "Battle.net 1.0 Classic";
    public const string Wc3ClassicSetName = "Warcraft III Classic";
    public const string Bnet2SetName = "Battle.net 2.0";

    /// <summary>The three bundled (not user-saved) icon sets offered wherever sets are picked: Manage Icons, the Bot menu, the Users list, Appearance.</summary>
    public static readonly IReadOnlyList<string> BundledNames = [Bnet1ClassicSetName, Wc3ClassicSetName, Bnet2SetName];

    /// <summary>
    /// "Battle.net 1.0 Classic" — the original classic.battle.net chat-icon set
    /// (classic.battle.net/info/icons.shtml), 28x14, sourced 2026-08-24. This is also what
    /// GameIconLoader's own bundled defaults are for every key here except blizz/sysop/mod-gavel/
    /// mega/ignore (which default to the sharper Warcraft III Classic art instead — see
    /// GameIconLoader's remarks) — selecting this set explicitly is how to get the small original
    /// look back for those five specifically. scbw ("W2sexp.png", the real StarCraft: Brood War
    /// badge — classic.battle.net's own /info/icons.shtml never actually distinguished it from
    /// plain StarCraft, a real bug fixed 2026-08-24 alongside Chat.ChatIcon.GetProductIconKey's
    /// matching PXES-mapping fix) and w3tft ("W2w3xp.png") both come from
    /// warcraft.wiki.gg/wiki/Warcraft_II_chat_icons instead — confirmed classic.battle.net
    /// itself never shipped one at all (its own site "never fully updated before it went down,"
    /// per the user); this wiki apparently caught a copy before that happened, filed oddly under
    /// the Warcraft II chat-icons page rather than Warcraft III's own. sc2 has no official
    /// classic-era icon at all (StarCraft II postdates this whole art style) — this one is
    /// user-made (2026-08-24, hand-drawn to match the aesthetic, dropped straight in at the same
    /// 28x14 size as the rest of this set) rather than sourced from anywhere. Not in Wc3ClassicSet
    /// below: that set is 64x64, and this icon is 28x14 — forcing it in would look inconsistently
    /// tiny/blurry next to the rest of that set. diablo-dot0..3 (the "how far have they gotten"
    /// badges GetProductIconKey picks for a Diablo user with no status icon — see its own remarks)
    /// are real Blizzard originals too, from classic.battle.net/info/icons.shtml's own "Diablo
    /// Icons" section (warrior.jpg/sorcerer.jpg/2dots.jpg/rogue.jpg for 0/1/2/3 dots respectively,
    /// sourced 2026-09-11) — contrary to the "went down" note above, that page is still live at
    /// its original URL as of this date, so it's worth re-checking directly for anything else
    /// still missing here rather than assuming it's gone. sc-stars0..10 and war2-axes0..8/
    /// war2-sword1/war2-sword2/war2-ranked are likewise real — not fabricated placeholders —
    /// extracted from the actual icons_STAR.bni and icons_W2BN.bni resource files themselves
    /// (bnetdocs.org/document/25/icons-bni documents the .bni binary format; the already-decoded
    /// individual PNGs live at files.bnetdocs.org/Battle.net/Icons/Extracted/STAR and .../W2BN —
    /// indices 8-17 of STAR and 8-18 of W2BN respectively, confirmed against the user's own
    /// firsthand knowledge of the exact win/rank thresholds since the .bni's icon-selection
    /// metadata isn't included in those extracted PNGs). guest.png (index 4, shared by both
    /// files) also came from here, replacing an earlier hand-drawn placeholder — see ChatIcon's
    /// remarks on GetStatusIconKey. Two ranked-plate variants these games' real icons.bni also
    /// define (a "top 5% of ladder" artsy background and a literal "#1 rank" plate) are NOT wired
    /// up anywhere despite being extracted (indices 20/21 of both files) — nothing in this app's
    /// data has ladder position/percentile, only the rating number itself, so there's no way to
    /// choose between them; every ranked user gets the plain "ranked" plate today.
    /// </summary>
    private static readonly (string Key, string Folder, string SourceKey)[] Bnet1ClassicSet =
    [
        ("sc", "GameIconsClassic", "sc"), ("scbw", "GameIconsClassic", "scbw"), ("jsc", "GameIconsClassic", "jsc"),
        ("sware", "GameIconsClassic", "sware"),
        ("sc-stars0", "GameIconsClassic", "sc-stars0"), ("sc-stars1", "GameIconsClassic", "sc-stars1"),
        ("sc-stars2", "GameIconsClassic", "sc-stars2"), ("sc-stars3", "GameIconsClassic", "sc-stars3"),
        ("sc-stars4", "GameIconsClassic", "sc-stars4"), ("sc-stars5", "GameIconsClassic", "sc-stars5"),
        ("sc-stars6", "GameIconsClassic", "sc-stars6"), ("sc-stars7", "GameIconsClassic", "sc-stars7"),
        ("sc-stars8", "GameIconsClassic", "sc-stars8"), ("sc-stars9", "GameIconsClassic", "sc-stars9"),
        ("sc-stars10", "GameIconsClassic", "sc-stars10"),
        ("war2", "GameIconsClassic", "war2"),
        ("war2-axes0", "GameIconsClassic", "war2-axes0"), ("war2-axes1", "GameIconsClassic", "war2-axes1"),
        ("war2-axes2", "GameIconsClassic", "war2-axes2"), ("war2-axes3", "GameIconsClassic", "war2-axes3"),
        ("war2-axes4", "GameIconsClassic", "war2-axes4"), ("war2-axes5", "GameIconsClassic", "war2-axes5"),
        ("war2-axes6", "GameIconsClassic", "war2-axes6"), ("war2-axes7", "GameIconsClassic", "war2-axes7"),
        ("war2-axes8", "GameIconsClassic", "war2-axes8"), ("war2-sword1", "GameIconsClassic", "war2-sword1"),
        ("war2-sword2", "GameIconsClassic", "war2-sword2"), ("war2-ranked", "GameIconsClassic", "war2-ranked"),
        ("war3", "GameIconsClassic", "war3"),
        ("war3-tier1", "GameIconsClassic", "war3-tier1"), ("war3-tier2-human", "GameIconsClassic", "war3-tier2-human"), ("war3-tier3-human", "GameIconsClassic", "war3-tier3-human"), ("war3-tier4-human", "GameIconsClassic", "war3-tier4-human"), ("war3-tier5-human", "GameIconsClassic", "war3-tier5-human"), ("war3-tier6-human", "GameIconsClassic", "war3-tier6-human"), ("war3-tier2-orc", "GameIconsClassic", "war3-tier2-orc"), ("war3-tier3-orc", "GameIconsClassic", "war3-tier3-orc"), ("war3-tier4-orc", "GameIconsClassic", "war3-tier4-orc"), ("war3-tier5-orc", "GameIconsClassic", "war3-tier5-orc"), ("war3-tier6-orc", "GameIconsClassic", "war3-tier6-orc"), ("war3-tier2-nightelf", "GameIconsClassic", "war3-tier2-nightelf"), ("war3-tier3-nightelf", "GameIconsClassic", "war3-tier3-nightelf"), ("war3-tier4-nightelf", "GameIconsClassic", "war3-tier4-nightelf"), ("war3-tier5-nightelf", "GameIconsClassic", "war3-tier5-nightelf"), ("war3-tier6-nightelf", "GameIconsClassic", "war3-tier6-nightelf"), ("war3-tier2-undead", "GameIconsClassic", "war3-tier2-undead"), ("war3-tier3-undead", "GameIconsClassic", "war3-tier3-undead"), ("war3-tier4-undead", "GameIconsClassic", "war3-tier4-undead"), ("war3-tier5-undead", "GameIconsClassic", "war3-tier5-undead"), ("war3-tier6-undead", "GameIconsClassic", "war3-tier6-undead"), ("war3-tier2-random", "GameIconsClassic", "war3-tier2-random"), ("war3-tier3-random", "GameIconsClassic", "war3-tier3-random"), ("war3-tier4-random", "GameIconsClassic", "war3-tier4-random"), ("war3-tier5-random", "GameIconsClassic", "war3-tier5-random"), ("war3-tier6-random", "GameIconsClassic", "war3-tier6-random"), ("war3-tier2-tourney", "GameIconsClassic", "war3-tier2-tourney"), ("war3-tier3-tourney", "GameIconsClassic", "war3-tier3-tourney"), ("war3-tier4-tourney", "GameIconsClassic", "war3-tier4-tourney"), ("war3-tier5-tourney", "GameIconsClassic", "war3-tier5-tourney"), ("war3-tier6-tourney", "GameIconsClassic", "war3-tier6-tourney"),
        ("w3tft", "GameIconsClassic", "w3tft"),
        ("diablo", "GameIconsClassic", "diablo"), ("dshr", "GameIconsClassic", "dshr"),
        ("diablo-dot0", "GameIconsClassic", "diablo-dot0"), ("diablo-dot1", "GameIconsClassic", "diablo-dot1"),
        ("diablo-dot2", "GameIconsClassic", "diablo-dot2"), ("diablo-dot3", "GameIconsClassic", "diablo-dot3"),
        ("diablo2", "GameIconsClassic", "diablo2"), ("d2exp", "GameIconsClassic", "d2exp"),
        ("chat", "GameIconsClassic", "chat"), ("blizz", "GameIconsClassic", "blizz"),
        ("sysop", "GameIconsClassic", "sysop"), ("mod-gavel", "GameIconsClassic", "mod-gavel"),
        ("mega", "GameIconsClassic", "mega"), ("guest", "GameIconsClassic", "guest"),
        ("ignore", "GameIconsClassic", "ignore"), ("sc2", "GameIconsClassic", "sc2"),
    ];

    /// <summary>
    /// "Warcraft III Classic" — the 64x64 set under Assets/GameIconsHD, sourced directly from
    /// classic.battle.net/war3/images/battle.net/icons/ (the WC3 ladder site's own icon folder —
    /// true Blizzard-hosted originals, not a third-party re-host, confirmed 2026-08-24 per the
    /// WC3 ladder icons.shtml page listing every file in that folder). sysop uses
    /// "bnet-battlenet.gif" (the real Admin icon, not a fallback), mega uses "bnet-speaker.gif"
    /// (the real Speaker icon) — both explicit user corrections replacing an earlier
    /// Battle.net-1.0-Classic-borrowed placeholder. jsc/sware/dshr still have no distinct HD art
    /// upstream, so they borrow their closest relative (StarCraft/StarCraft/Diablo); war3/w3tft
    /// aren't in that same folder either (it's a ladder-status icon set, not per-product game
    /// icons) and stay sourced from wowpedia.fandom.com/wiki/Warcraft_III_chat_icons instead. sc2
    /// is user-made (2026-08-24, "bnet-sc2_war3_style.png") — StarCraft II postdates this whole
    /// art style, so there's no official upstream icon.
    /// </summary>
    private static readonly (string Key, string Folder, string SourceKey)[] Wc3ClassicSet =
    [
        ("sc", "GameIconsHD", "sc"), ("scbw", "GameIconsHD", "scbw"), ("jsc", "GameIconsHD", "sc"),
        ("sware", "GameIconsHD", "sc"),
        ("sc-stars0", "GameIconsHD", "sc-stars0"), ("sc-stars1", "GameIconsHD", "sc-stars1"),
        ("sc-stars2", "GameIconsHD", "sc-stars2"), ("sc-stars3", "GameIconsHD", "sc-stars3"),
        ("sc-stars4", "GameIconsHD", "sc-stars4"), ("sc-stars5", "GameIconsHD", "sc-stars5"),
        ("sc-stars6", "GameIconsHD", "sc-stars6"), ("sc-stars7", "GameIconsHD", "sc-stars7"),
        ("sc-stars8", "GameIconsHD", "sc-stars8"), ("sc-stars9", "GameIconsHD", "sc-stars9"),
        ("sc-stars10", "GameIconsHD", "sc-stars10"),
        ("war2", "GameIconsHD", "war2"),
        ("war2-axes0", "GameIconsHD", "war2-axes0"), ("war2-axes1", "GameIconsHD", "war2-axes1"),
        ("war2-axes2", "GameIconsHD", "war2-axes2"), ("war2-axes3", "GameIconsHD", "war2-axes3"),
        ("war2-axes4", "GameIconsHD", "war2-axes4"), ("war2-axes5", "GameIconsHD", "war2-axes5"),
        ("war2-axes6", "GameIconsHD", "war2-axes6"), ("war2-axes7", "GameIconsHD", "war2-axes7"),
        ("war2-axes8", "GameIconsHD", "war2-axes8"), ("war2-sword1", "GameIconsHD", "war2-sword1"),
        ("war2-sword2", "GameIconsHD", "war2-sword2"), ("war2-ranked", "GameIconsHD", "war2-ranked"),
        ("war3", "GameIconsHD", "war3"),
        ("war3-tier1", "GameIconsHD", "war3-tier1"), ("war3-tier2-human", "GameIconsHD", "war3-tier2-human"), ("war3-tier3-human", "GameIconsHD", "war3-tier3-human"), ("war3-tier4-human", "GameIconsHD", "war3-tier4-human"), ("war3-tier5-human", "GameIconsHD", "war3-tier5-human"), ("war3-tier6-human", "GameIconsHD", "war3-tier6-human"), ("war3-tier2-orc", "GameIconsHD", "war3-tier2-orc"), ("war3-tier3-orc", "GameIconsHD", "war3-tier3-orc"), ("war3-tier4-orc", "GameIconsHD", "war3-tier4-orc"), ("war3-tier5-orc", "GameIconsHD", "war3-tier5-orc"), ("war3-tier6-orc", "GameIconsHD", "war3-tier6-orc"), ("war3-tier2-nightelf", "GameIconsHD", "war3-tier2-nightelf"), ("war3-tier3-nightelf", "GameIconsHD", "war3-tier3-nightelf"), ("war3-tier4-nightelf", "GameIconsHD", "war3-tier4-nightelf"), ("war3-tier5-nightelf", "GameIconsHD", "war3-tier5-nightelf"), ("war3-tier6-nightelf", "GameIconsHD", "war3-tier6-nightelf"), ("war3-tier2-undead", "GameIconsHD", "war3-tier2-undead"), ("war3-tier3-undead", "GameIconsHD", "war3-tier3-undead"), ("war3-tier4-undead", "GameIconsHD", "war3-tier4-undead"), ("war3-tier5-undead", "GameIconsHD", "war3-tier5-undead"), ("war3-tier6-undead", "GameIconsHD", "war3-tier6-undead"), ("war3-tier2-random", "GameIconsHD", "war3-tier2-random"), ("war3-tier3-random", "GameIconsHD", "war3-tier3-random"), ("war3-tier4-random", "GameIconsHD", "war3-tier4-random"), ("war3-tier5-random", "GameIconsHD", "war3-tier5-random"), ("war3-tier6-random", "GameIconsHD", "war3-tier6-random"), ("war3-tier2-tourney", "GameIconsHD", "war3-tier2-tourney"), ("war3-tier3-tourney", "GameIconsHD", "war3-tier3-tourney"), ("war3-tier4-tourney", "GameIconsHD", "war3-tier4-tourney"), ("war3-tier5-tourney", "GameIconsHD", "war3-tier5-tourney"), ("war3-tier6-tourney", "GameIconsHD", "war3-tier6-tourney"),
        ("w3tft", "GameIconsHD", "w3tft"), ("diablo", "GameIconsHD", "diablo"), ("dshr", "GameIconsHD", "diablo"),
        ("diablo-dot0", "GameIconsHD", "diablo-dot0"), ("diablo-dot1", "GameIconsHD", "diablo-dot1"),
        ("diablo-dot2", "GameIconsHD", "diablo-dot2"), ("diablo-dot3", "GameIconsHD", "diablo-dot3"),
        ("diablo2", "GameIconsHD", "diablo2"), ("d2exp", "GameIconsHD", "d2exp"), ("chat", "GameIconsHD", "chat"),
        ("blizz", "GameIconsHD", "blizz"), ("sysop", "GameIconsHD", "sysop"),
        ("mod-gavel", "GameIconsHD", "mod-gavel"), ("mega", "GameIconsHD", "mega"),
        ("guest", "GameIconsHD", "guest"), ("ignore", "GameIconsHD", "ignore"),
        ("sc2", "GameIconsHD", "sc2"),
    ];

    /// <summary>
    /// "Battle.net 2.0" — the official account.battle.net game-icon set (rasterized SVGs under
    /// Assets/GameIconsBnet2), plus the Warcraft III Classic set's status badges per explicit
    /// request ("for Battle.net 2.0 we'll want to use the status badges from War3 set") — with one
    /// exception: mod-gavel uses "mod-gavel-glow.png", a green-glow variant of the same War3
    /// hammer (generated 2026-08-24 via a blurred green-tinted copy of the icon's own silhouette
    /// drawn behind the crisp original — "a gentle green glow around it," per request) rather than
    /// the plain one, to read as more at-home next to Bnet2's brighter modern art. sc2 uses the
    /// real official account.battle.net StarCraft II SVG (account.battle.net/static/images/
    /// game-icons/starcraft-ii.svg) — fixes a real bug where it had no entry in this set at all,
    /// so switching to Battle.net 2.0 after Battle.net 1.0 Classic/Warcraft III Classic left
    /// whichever classic-style sc2 icon was applied stuck in place instead of reverting. Several
    /// keys intentionally share one source image, matching how Blizzard's own modern branding
    /// doesn't distinguish them: StarCraft/Brood War/the Japanese release/the shareware trial all
    /// point at one "StarCraft: Remastered" icon, same idea for Warcraft III/TFT and Diablo II/
    /// Lord of Destruction.
    /// </summary>
    private static readonly (string Key, string Folder, string SourceKey)[] Bnet2Set =
    [
        ("diablo2", "GameIconsBnet2", "diablo-ii"),
        ("d2exp", "GameIconsBnet2", "diablo-ii"),
        ("war3", "GameIconsBnet2", "warcraft-iii"),
        ("w3tft", "GameIconsBnet2", "warcraft-iii"),
        ("war2", "GameIconsBnet2", "warcraft-ii-remastered"),
        ("sc", "GameIconsBnet2", "starcraft-remastered"),
        ("scbw", "GameIconsBnet2", "starcraft-remastered"),
        ("jsc", "GameIconsBnet2", "starcraft-remastered"),
        ("sware", "GameIconsBnet2", "starcraft-remastered"),
        ("sc2", "GameIconsBnet2", "starcraft-ii"),
        ("blizz", "GameIconsHD", "blizz"),
        ("sysop", "GameIconsHD", "sysop"),
        ("mod-gavel", "GameIconsBnet2", "mod-gavel-glow"),
        ("mega", "GameIconsHD", "mega"),
        ("guest", "GameIconsHD", "guest"),
        ("ignore", "GameIconsHD", "ignore"),
        ("diablo", "GameIconsBnet2", "diablo"),
        ("dshr", "GameIconsBnet2", "diablo"),
        ("diablo-dot0", "GameIconsHD", "diablo-dot0"),
        ("diablo-dot1", "GameIconsHD", "diablo-dot1"),
        ("diablo-dot2", "GameIconsHD", "diablo-dot2"),
        ("diablo-dot3", "GameIconsHD", "diablo-dot3"),
        ("sc-stars0", "GameIconsHD", "sc-stars0"), ("sc-stars1", "GameIconsHD", "sc-stars1"),
        ("sc-stars2", "GameIconsHD", "sc-stars2"), ("sc-stars3", "GameIconsHD", "sc-stars3"),
        ("sc-stars4", "GameIconsHD", "sc-stars4"), ("sc-stars5", "GameIconsHD", "sc-stars5"),
        ("sc-stars6", "GameIconsHD", "sc-stars6"), ("sc-stars7", "GameIconsHD", "sc-stars7"),
        ("sc-stars8", "GameIconsHD", "sc-stars8"), ("sc-stars9", "GameIconsHD", "sc-stars9"),
        ("sc-stars10", "GameIconsHD", "sc-stars10"),
        ("war2-axes0", "GameIconsHD", "war2-axes0"), ("war2-axes1", "GameIconsHD", "war2-axes1"),
        ("war2-axes2", "GameIconsHD", "war2-axes2"), ("war2-axes3", "GameIconsHD", "war2-axes3"),
        ("war2-axes4", "GameIconsHD", "war2-axes4"), ("war2-axes5", "GameIconsHD", "war2-axes5"),
        ("war2-axes6", "GameIconsHD", "war2-axes6"), ("war2-axes7", "GameIconsHD", "war2-axes7"),
        ("war2-axes8", "GameIconsHD", "war2-axes8"), ("war2-sword1", "GameIconsHD", "war2-sword1"),
        ("war2-sword2", "GameIconsHD", "war2-sword2"), ("war2-ranked", "GameIconsHD", "war2-ranked"),
        ("war3-tier1", "GameIconsHD", "war3-tier1"), ("war3-tier2-human", "GameIconsHD", "war3-tier2-human"), ("war3-tier3-human", "GameIconsHD", "war3-tier3-human"), ("war3-tier4-human", "GameIconsHD", "war3-tier4-human"), ("war3-tier5-human", "GameIconsHD", "war3-tier5-human"), ("war3-tier6-human", "GameIconsHD", "war3-tier6-human"), ("war3-tier2-orc", "GameIconsHD", "war3-tier2-orc"), ("war3-tier3-orc", "GameIconsHD", "war3-tier3-orc"), ("war3-tier4-orc", "GameIconsHD", "war3-tier4-orc"), ("war3-tier5-orc", "GameIconsHD", "war3-tier5-orc"), ("war3-tier6-orc", "GameIconsHD", "war3-tier6-orc"), ("war3-tier2-nightelf", "GameIconsHD", "war3-tier2-nightelf"), ("war3-tier3-nightelf", "GameIconsHD", "war3-tier3-nightelf"), ("war3-tier4-nightelf", "GameIconsHD", "war3-tier4-nightelf"), ("war3-tier5-nightelf", "GameIconsHD", "war3-tier5-nightelf"), ("war3-tier6-nightelf", "GameIconsHD", "war3-tier6-nightelf"), ("war3-tier2-undead", "GameIconsHD", "war3-tier2-undead"), ("war3-tier3-undead", "GameIconsHD", "war3-tier3-undead"), ("war3-tier4-undead", "GameIconsHD", "war3-tier4-undead"), ("war3-tier5-undead", "GameIconsHD", "war3-tier5-undead"), ("war3-tier6-undead", "GameIconsHD", "war3-tier6-undead"), ("war3-tier2-random", "GameIconsHD", "war3-tier2-random"), ("war3-tier3-random", "GameIconsHD", "war3-tier3-random"), ("war3-tier4-random", "GameIconsHD", "war3-tier4-random"), ("war3-tier5-random", "GameIconsHD", "war3-tier5-random"), ("war3-tier6-random", "GameIconsHD", "war3-tier6-random"), ("war3-tier2-tourney", "GameIconsHD", "war3-tier2-tourney"), ("war3-tier3-tourney", "GameIconsHD", "war3-tier3-tourney"), ("war3-tier4-tourney", "GameIconsHD", "war3-tier4-tourney"), ("war3-tier5-tourney", "GameIconsHD", "war3-tier5-tourney"), ("war3-tier6-tourney", "GameIconsHD", "war3-tier6-tourney"),
    ];

    private static readonly Dictionary<string, Avalonia.Media.Imaging.Bitmap?> Bnet2Cache = [];

    /// <summary>
    /// A key's icon from the Battle.net 2.0 set whatever set is showing, for the Battle.net 2.0
    /// friends list. Keys that set doesn't cover (d4, bnet2, offline...) come as usual.
    /// </summary>
    public static Avalonia.Media.Imaging.Bitmap? GetBnet2(string key)
    {
        if (Bnet2Set.FirstOrDefault(entry => entry.Key == key) is not { Folder: not null } entry)
        {
            return GameIconLoader.Get(key);
        }

        if (!Bnet2Cache.TryGetValue(key, out var bitmap))
        {
            using var stream = AssetLoader.Open(new Uri($"avares://Invigoration.App/Assets/{entry.Folder}/{entry.SourceKey}.png"));
            Bnet2Cache[key] = bitmap = new Avalonia.Media.Imaging.Bitmap(stream);
        }

        return bitmap;
    }

    /// <summary>The bundled sets, then every saved one whose name doesn't clash with them.</summary>
    public static IReadOnlyList<string> All() =>
        BundledNames.Concat(IconSetStore.ListSets().Where(name => !IsBundled(name))).ToList();

    public static bool IsBundled(string name) => BundledNames.Contains(name);

    /// <summary>Applies <paramref name="name"/>, whether bundled or saved, and records it as the set showing.</summary>
    public static void Apply(string name)
    {
        var mapping = name switch
        {
            Bnet1ClassicSetName => Bnet1ClassicSet,
            Wc3ClassicSetName => Wc3ClassicSet,
            Bnet2SetName => Bnet2Set,
            _ => null,
        };

        if (mapping is null)
        {
            IconSetStore.ApplySet(name);
            return;
        }

        foreach (var (key, folder, sourceKey) in mapping)
        {
            using var stream = AssetLoader.Open(new Uri($"avares://Invigoration.App/Assets/{folder}/{sourceKey}.png"));
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            IconOverrideStore.SetOverrideBytes(key, buffer.ToArray(), ".png");
        }

        IconSetStore.ActiveSetName = name;
    }

    /// <summary>
    /// Applies <paramref name="name"/> unless it's the set already showing, which would only copy the
    /// same images in again and undo any single icon changed in Manage Icons since. "" (a bot with
    /// no set of its own) leaves whatever is showing.
    /// </summary>
    public static void ApplyIfNotShowing(string name)
    {
        if (!string.IsNullOrEmpty(name) && name != IconSetStore.ActiveSetName)
        {
            Apply(name);
        }
    }
}
