namespace Invigoration.Core.Protocol;

/// <summary>
/// Where a product/connection-mode combination is actually reachable today.
/// Several classic products were retired from Blizzard's live service but
/// still work fine against PVPGN/private servers, so this is tracked per
/// product rather than assumed.
/// </summary>
public enum ServerCompatibility
{
    /// <summary>Works against official Battle.net and PVPGN/private servers.</summary>
    Both,

    /// <summary>Retired from official Battle.net — only private/PVPGN servers will accept it.</summary>
    PrivateOnly,
}

public sealed record BncsProductInfo(
    string WireCode,
    string DisplayName,
    byte? BnlsProductByte,
    ServerCompatibility Compatibility,
    string? Notes = null,
    string? IconKeyOverride = null,
    bool SupportsFriendsList = true)
{
    /// <summary>
    /// Icon key for a *known, exact* product — used by the game-selection
    /// picker, where there's no ambiguity about which product this is.
    /// Defaults to <see cref="Chat.ChatIcon.GetProductIconKey"/>'s
    /// statstring-based mapping, which deliberately merges StarCraft/Brood
    /// War and Warcraft III/TFT onto one icon each (the wire statstring
    /// alone can't always tell them apart, e.g. StarCraft: Remastered
    /// self-identifies as plain Brood War) — <see cref="IconKeyOverride"/>
    /// exists so the picker can still show each of those a distinct icon
    /// when the product is explicitly chosen rather than inferred.
    /// </summary>
    public string IconKey => IconKeyOverride ?? Chat.ChatIcon.GetProductIconKey(WireCode);
}

/// <summary>
/// Catalog of legacy-BNCS products this engine speaks, keyed by their 4-character
/// wire product code (stored in wire/reversed form throughout this codebase,
/// e.g. "VD2D" is the wire form of the human-readable "D2DV").
///
/// Scope note: StarCraft II and StarCraft: Remastered's *modern* Battle.net
/// backend (protobuf/WebSocket + bit-packed RC4 TCP, per superioritybot.com/PROTOCOL)
/// is an entirely separate protocol from BNCS and is intentionally out of scope
/// here — planned as a future, separate engine/project. StarCraft: Remastered's
/// classic-chat connectivity (it reportedly still joins the same chat servers as
/// D2DV/D2XP/W2BN) does fall inside BNCS's scope, but its exact EXE
/// version/hash data for the BNLS version-check differs from the original 1998
/// client and hasn't been verified yet — tracked as a follow-up, not implemented
/// here. Until then, StarCraft (RATS) / Brood War (PXES) are treated as
/// PVPGN-only, matching the original 1998 client's current (retired) status on
/// official Battle.net.
/// </summary>
public static class BncsProduct
{
    public const string DiabloII = "VD2D";
    public const string DiabloIILoD = "PX2D";
    public const string Warcraft2BNE = "NB2W";
    public const string Warcraft3 = "3RAW";
    public const string Warcraft3TFT = "PX3W";
    public const string Starcraft = "RATS";
    public const string StarcraftJapanese = "RTSJ";
    public const string StarcraftBroodWar = "PXES";
    public const string Diablo = "LTRD";

    public static readonly IReadOnlyDictionary<string, BncsProductInfo> Catalog =
        new Dictionary<string, BncsProductInfo>
        {
            [DiabloII] = new(DiabloII, "Diablo II", 0x4, ServerCompatibility.Both),
            [DiabloIILoD] = new(DiabloIILoD, "Diablo II: Lord of Destruction", 0x5, ServerCompatibility.Both),
            [Warcraft2BNE] = new(
                Warcraft2BNE,
                "Warcraft II: Battle.net Edition",
                0x3,
                ServerCompatibility.Both),
            [Diablo] = new(
                Diablo,
                "Diablo",
                0x9,
                ServerCompatibility.Both,
                "Official Battle.net restricts Diablo (1) to Public chat channels only; PVPGN has no such restriction.",
                SupportsFriendsList: false),
            [Warcraft3] = new(
                Warcraft3,
                "Warcraft III: Reign of Chaos",
                0x7,
                ServerCompatibility.PrivateOnly,
                "Retired from official Battle.net entirely; only usable against PVPGN."),
            [Warcraft3TFT] = new(
                Warcraft3TFT,
                "Warcraft III: The Frozen Throne",
                0x8,
                ServerCompatibility.PrivateOnly,
                "Retired from official Battle.net entirely; only usable against PVPGN.",
                IconKeyOverride: "w3tft"),
            [Starcraft] = new(
                Starcraft,
                "StarCraft (1998)",
                0x1,
                ServerCompatibility.PrivateOnly,
                "Original client retired from official Battle.net; only usable against PVPGN. " +
                "StarCraft: Remastered reportedly reuses this product code on live Battle.net with " +
                "different EXE version/hash data — not yet implemented, see class remarks."),
            [StarcraftBroodWar] = new(
                StarcraftBroodWar,
                "StarCraft: Brood War (1998)",
                0x2,
                ServerCompatibility.PrivateOnly,
                "Original client retired from official Battle.net; only usable against PVPGN.",
                IconKeyOverride: "scbw"),
            [StarcraftJapanese] = new(
                StarcraftJapanese,
                "StarCraft (Japanese release)",
                0x6,
                ServerCompatibility.PrivateOnly,
                "Retired from official Battle.net; only usable against PVPGN."),
        };

    /// <summary>Official Battle.net regional chat servers, offered only for products where <see cref="ServerCompatibility.Both"/> applies.</summary>
    public static readonly IReadOnlyList<string> OfficialBattlenetServers =
    [
        "uswest.battle.net",
        "useast.battle.net",
        "asia.battle.net",
        "europe.battle.net",
    ];

    /// <summary>Well-known public PVPGN-family servers, offered as quick picks alongside a free-text custom server field.</summary>
    public static readonly IReadOnlyList<string> SuggestedPrivateServers =
    [
        "atlas.bnetdocs.org",
        "pvpgn.bnetdocs.org",
    ];

    public static byte? GetBnlsProductByte(string wireCode) =>
        Catalog.TryGetValue(wireCode, out var info) ? info.BnlsProductByte : null;

    public static string GetDisplayName(string wireCode) =>
        Catalog.TryGetValue(wireCode, out var info) ? info.DisplayName : wireCode;

    /// <summary>Products that use the NLS/SRP-based new login system instead of the old double-hash system.</summary>
    public static bool UsesNewLoginSystem(string wireCode) => wireCode is Warcraft3 or Warcraft3TFT;

    /// <summary>Expansion products that authenticate with a classic+expansion CD-key pair.</summary>
    public static bool RequiresExpansionCdKey(string wireCode) => wireCode is DiabloIILoD or Warcraft3TFT;

    /// <summary>Whether this product needs a CD-key at all — Diablo (1) doesn't.</summary>
    public static bool RequiresCdKey(string wireCode) => wireCode != Diablo;

    public static ServerCompatibility GetServerCompatibility(string wireCode) =>
        Catalog.TryGetValue(wireCode, out var info) ? info.Compatibility : ServerCompatibility.Both;

    /// <summary>
    /// Whether this product's server pushes SID_FRIENDSLIST/UPDATE/ADD/REMOVE/POSITION
    /// at all — false only for Diablo (1), which predates the Friends
    /// feature entirely and has no client-side "/f" command for it either.
    /// Warcraft II: Battle.net Edition does support it (its client has a
    /// "/f list" command) despite bnetdocs.org's "Used By" list omitting it.
    /// </summary>
    public static bool SupportsFriendsList(string wireCode) =>
        Catalog.TryGetValue(wireCode, out var info) ? info.SupportsFriendsList : true;

    /// <summary>Heuristic check used to warn before connecting a PrivateOnly product to what looks like official Battle.net.</summary>
    public static bool LooksLikeOfficialBattlenetHost(string hostOrAddress) =>
        hostOrAddress.Equals("battle.net", StringComparison.OrdinalIgnoreCase) ||
        hostOrAddress.EndsWith(".battle.net", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// True if this connection mode/product combination is known to be
    /// blocked by the target server type. This is advisory (the server is the
    /// real authority) — intended to warn the user before they waste a
    /// connection attempt, not to enforce anything.
    /// </summary>
    public static bool IsLikelyIncompatible(string wireCode, string hostOrAddress) =>
        GetServerCompatibility(wireCode) == ServerCompatibility.PrivateOnly &&
        LooksLikeOfficialBattlenetHost(hostOrAddress);
}
