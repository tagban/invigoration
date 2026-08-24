using System.Text.Json;
using System.Text.Json.Serialization;
using Invigoration.Core.Chat;
using Invigoration.Core.Networking;
using Stimpak;

namespace Invigoration.Core.Config;

/// <summary>Per-bot settings. Replaces the VB6 BotData/BotNetData globals.</summary>
public sealed class BotConfig
{
    private static readonly JsonSerializerOptions CloneOptions = new()
    {
        Converters = { new JsonStringEnumConverter() },
    };

    /// <summary>
    /// Deep-clones a config (JSON round-trip) so it can be edited safely
    /// without mutating the original until the edit is explicitly saved.
    /// </summary>
    public static BotConfig Clone(BotConfig source) =>
        JsonSerializer.Deserialize<BotConfig>(JsonSerializer.Serialize(source, CloneOptions), CloneOptions)!;


    public string DisplayName { get; set; } = "New Bot";
    public string Username { get; set; } = "";

    /// <summary>
    /// Obfuscated at rest (see <see cref="ObfuscatedPasswordJsonConverter"/>)
    /// — always plaintext in memory. Stored on disk as "[base64]"; a
    /// plaintext value typed directly into bots.json by hand (no brackets)
    /// still loads and works, and gets wrapped again on the next save.
    /// </summary>
    [JsonConverter(typeof(ObfuscatedPasswordJsonConverter))]
    public string Password { get; set; } = "";

    // --- StarCraft II (modern Battle.net login, separate from the classic BNCS fields above) ---

    /// <summary>
    /// Which BattlenetCredentialProfileStore profile this bot's Battle.net
    /// login comes from — "" means none assigned yet (a bot config predating
    /// this feature, or one whose assigned profile was since deleted);
    /// BotEngine.Sc2.cs auto-creates and stamps one in on first connect.
    /// Sharing the same Id across two bots (e.g. an SC2 bot and a future
    /// WC3:Reforged bot on the same account) makes them share one signed-in
    /// session instead of prompting for a separate login each.
    /// </summary>
    public string BattlenetCredentialProfileId { get; set; } = "";

    /// <summary>
    /// The channels joined last time this bot was connected (SC2/SC:R/WC3:R only — classic
    /// BNCS/Chat-Telnet are single-channel and have nothing to remember here). Handed straight
    /// to Stimpak's own StimpakConnectOptions.Channels on the next connect, which restores them
    /// natively (an empty list just means the default "General") — not replayed by hand after
    /// the fact. Kept in sync automatically as channels are joined/left — not meant to be
    /// hand-edited. Stimpak's ChannelTarget is reused directly here (it already serializes via
    /// the same JsonPolymorphic/JsonDerivedType attributes ConfigStore's plain
    /// System.Text.Json options already understand) rather than re-deriving an equivalent type.
    /// </summary>
    public List<ChannelTarget> Sc2LastChannels { get; set; } = new();

    /// <summary>
    /// Which named group this bot's top-level tab belongs to in MainWindow — "" means
    /// ungrouped, shown as its own individual tab like before. Purely a display grouping (e.g.
    /// several bots on the same server), set from the Config window; has no effect on
    /// connection behavior.
    /// </summary>
    public string TabGroup { get; set; } = "";

    public string CdKey { get; set; } = "";

    /// <summary>
    /// Second CD-key required by expansion products (D2:LoD, WC3:TFT) —
    /// see <see cref="Protocol.BncsProduct.RequiresExpansionCdKey"/>. Unused
    /// for single-key products.
    /// </summary>
    public string ExpansionCdKey { get; set; } = "";

    public string BotMaster { get; set; } = "";
    public string Trigger { get; set; } = "!";
    public bool UseUdp { get; set; }
    public string BattlenetServer { get; set; } = "useast.battle.net";
    public int BattlenetPort { get; set; } = 6112;
    public string BnlsServer { get; set; } = "bnls.bnetdocs.org";
    public int BnlsPort { get; set; } = 9367;

    /// <summary>
    /// Manual override for the version byte BNLS_REQUESTVERSIONBYTE normally
    /// supplies (see Auth.AuthState.VersionByte) — empty means use whatever
    /// BNLS returns, which is the normal/default behavior. Since this bot has
    /// no local hashing to fall back on, BNLS is otherwise the only source
    /// for this value — some private/PVPGN servers are pinned to an older
    /// game-client version than BNLS's own database assumes for a product,
    /// so their SID_AUTH_CHECK rejects the byte BNLS hands back even though
    /// the CD-key/hash portions are fine. Accepts a "0x"-prefixed hex value
    /// (e.g. "0x1A") or plain decimal; anything else is treated as unset.
    /// </summary>
    public string VersionByteOverride { get; set; } = "";

    public string HomeChannel { get; set; } = "";

    /// <summary>When true, this bot connects automatically when the app starts, instead of waiting for a manual Connect click.</summary>
    public bool AutoConnectOnStartup { get; set; }

    /// <summary>
    /// When true, an unexpected disconnect (never one caused by clicking
    /// Disconnect) triggers an automatic reconnect after AutoReconnectDelaySeconds.
    /// Off by default — see BotEngine.MaybeScheduleAutoReconnect's remarks:
    /// on official Battle.net specifically, a dropped connection can be the
    /// server enforcing SID_REQUIREDWORK/ExtraWork compliance (bot
    /// detection) this client doesn't implement, and repeated automatic
    /// reconnects against that could look more automated, not less.
    /// </summary>
    public bool AutoReconnect { get; set; }

    public int AutoReconnectDelaySeconds { get; set; } = 20;

    /// <summary>
    /// Routes every connection this bot makes (BNCS, BNLS, D2 realm) through
    /// a proxy — the only client-side lever against a third-party server's
    /// per-IP connection or flood limits, since several of the user's own
    /// bots (even individually well-behaved) can still look like a burst
    /// from one IP. Off by default, no cost when disabled.
    /// </summary>
    public bool ProxyEnabled { get; set; }

    public ProxyProtocol ProxyProtocol { get; set; } = ProxyProtocol.Socks5;

    public string ProxyHost { get; set; } = "";

    public int ProxyPort { get; set; } = 1080;

    public string ProxyUsername { get; set; } = "";

    /// <summary>Obfuscated at rest like Password — see ObfuscatedPasswordJsonConverter.</summary>
    [JsonConverter(typeof(ObfuscatedPasswordJsonConverter))]
    public string ProxyPassword { get; set; } = "";

    /// <summary>
    /// Turns on this bot's participation in the clan-management feature (the
    /// roster/rank/alias/trivia-score system in <see cref="Clan.ClanRosterStore"/>,
    /// not Battle.net's own in-game clan protocol) — off by default, opt-in
    /// per bot like BNU`Bot's plugin toggles, since not every bot is meant to
    /// run clan commands. The roster itself is shared across every bot
    /// regardless of platform (Classic BNCS today; SC2/SC:R once their native
    /// chat layers exist), so turning this on for bots on different games
    /// still has them all reading/writing the same clan.
    /// </summary>
    public bool ClanFeatureEnabled { get; set; }

    /// <summary>
    /// Rank auto-assigned to anyone who talks/emotes/whispers and isn't
    /// already a tracked clan-roster member (only when ClanFeatureEnabled is
    /// on) — builds an ongoing roster of everyone seen, not just people
    /// explicitly added, so ranks can be handed out later. Empty disables
    /// auto-registration; existing members keep whatever rank they already have.
    /// </summary>
    public string DefaultRank { get; set; } = "Trivia Participant";

    /// <summary>A member whose Rank matches this (case-insensitive) is blocked from every bot command, including the otherwise-always-open trivia "join"/"score" — set a member's rank to this via "clanrank" to revoke access.</summary>
    public string BannedRank { get; set; } = "Banned";

    /// <summary>Turns the trivia game on/off for this bot. When false, every "trivia" command (on/off/score/join) is unavailable — matches ClanFeatureEnabled's per-bot opt-in pattern. True by default since trivia has no other gate of its own.</summary>
    public bool TriviaFeatureEnabled { get; set; } = true;

    /// <summary>
    /// When non-empty, this bot shares its trivia round with every other bot
    /// (any game) naming the same group — one bot's "!trivia on" relays the
    /// question to every group member's own channel, and a correct answer on
    /// any of their channels ends the question for the whole group. Lets
    /// e.g. a Warcraft II bot and a StarCraft II bot, run by the same
    /// person, host one shared trivia game across both. Empty means this bot
    /// only ever runs its own independent round.
    /// </summary>
    public string TriviaGroup { get; set; } = "";

    /// <summary>Points awarded for answering correctly before the first hint is shown (i.e. within the first ~10 seconds). The richest of the three tiers, since no hint was needed.</summary>
    public double TriviaPointsBeforeFirstHint { get; set; } = 1.25;

    /// <summary>Points awarded for answering correctly after the first hint but before the second (~10-20 seconds in).</summary>
    public double TriviaPointsAfterFirstHint { get; set; } = 1.0;

    /// <summary>Points awarded for answering correctly after the second hint (~20-30 seconds in, right before time runs out).</summary>
    public double TriviaPointsAfterSecondHint { get; set; } = 0.75;

    /// <summary>
    /// Minimum delay between outgoing chat messages, in milliseconds — shared
    /// process-wide across every bot the user runs, not just this one, since
    /// a per-IP flood detector can still see multiple politely-spaced bots as
    /// a burst. So a burst (e.g. trivia asking a question then almost
    /// immediately announcing a fast correct answer, especially across two
    /// linked bots) doesn't trip the server's flood protection and get
    /// disconnected/banned. Raise it if you still get flooded off; lower it
    /// if your server is more permissive.
    /// </summary>
    public int FloodProtectionDelayMs { get; set; } = 2000;

    /// <summary>
    /// When on, Join/Leave lines never show in this bot's chat log at all — unconditional,
    /// unlike HideJoinLeaveSpamEnabled below which only hides them once a user's rate crosses a
    /// threshold. Same "display filter only" guarantee: roster tracking, rank behaviors, and
    /// JoinCount all still happen exactly as normal underneath.
    /// </summary>
    public bool SuppressJoinLeaveNotifications { get; set; }

    /// <summary>
    /// When on, a user who racks up more than HideJoinLeaveSpamThreshold
    /// Join/Leave events within a rolling HideJoinLeaveSpamWindowSeconds
    /// window stops having further Join/Leave lines shown in the chat log
    /// until they quiet back down (the window ages out) — a flaky connection
    /// bouncing in and out doesn't spam the log. Purely a display filter:
    /// roster tracking, rank behaviors (auto-whisper/kick/ban), and
    /// JoinCount all still happen exactly as normal underneath, nothing
    /// about the actual event handling changes, only whether that one line
    /// gets written to the visible chat log. On by default — there's no
    /// functional downside to filtering pure log noise, so this should just
    /// work without needing to be found and turned on first.
    /// </summary>
    public bool HideJoinLeaveSpamEnabled { get; set; } = true;

    public int HideJoinLeaveSpamThreshold { get; set; } = 3;

    public int HideJoinLeaveSpamWindowSeconds { get; set; } = 60;

    /// <summary>
    /// When on, a small game/client icon is shown next to a speaker's name on their chat/emote
    /// lines — the same icon key the userlist/friends list already derive from a classic BNCS
    /// user's statstring (Chat.ChatIcon.GetProductIconKey), or this bot's own product icon for
    /// every speaker on a Stimpak-backed (SC2/SC:R/WC3:R) bot, since Stimpak's roster data has no
    /// per-user product field to distinguish speakers by (same limitation noted on the friends
    /// list). Off by default — some noise for a feature not everyone wants in a dense chat log.
    /// </summary>
    public bool ShowUserIconsInChat { get; set; }

    /// <summary>4-character BNCS product code, e.g. "VD2D" = Diablo II, "PX2D" = Diablo II: LoD.</summary>
    public string Product { get; set; } = "VD2D";

    public string Realm { get; set; } = "";
    public bool ZeroPing { get; set; }
    public bool NegPing { get; set; }

    /// <summary>
    /// Which saved <see cref="Config.IconSetStore"/> set this bot's icons come
    /// from — "" means the bundled defaults (plus any global overrides set
    /// via the old single-active-set model). Applied by swapping in that
    /// set's files as the active <see cref="IconOverrideStore"/> overrides
    /// whenever this bot's tab becomes selected (see MainWindowViewModel);
    /// icon lookup itself has no per-bot concept, it's whichever set was
    /// last swapped in.
    /// </summary>
    public string IconSetName { get; set; } = "";

    /// <summary>Which named color set this bot's chat log and color codes render with.</summary>
    public ChatColorScheme ChatColorScheme { get; set; } = ChatColorScheme.Invigoration;

    /// <summary>User-edited palette, used when <see cref="ChatColorScheme"/> is <see cref="Chat.ChatColorScheme.Custom"/>. Defaults seeded from the Invigoration scheme.</summary>
    public CustomChatPalette CustomColors { get; set; } = new();

    /// <summary>Display name for <see cref="CustomColors"/> — shown in the editor and used as the default filename/label when exporting to share with others.</summary>
    public string CustomColorSchemeName { get; set; } = "My Color Scheme";

    public DiscordBridgeConfig Discord { get; set; } = new();
}

/// <summary>
/// Every ChatPalette role as a 0xRRGGBB packed int, editable one at a time
/// from the config window's color pickers. Defaults match ChatPalette.Invigoration.
/// </summary>
public sealed class CustomChatPalette
{
    public int Background { get; set; } = 0x242424;
    public int White { get; set; } = 0xFFFFFF;
    public int Channel { get; set; } = 0x00CE00;
    public int Info { get; set; } = 0x00C0C0;
    public int Error { get; set; } = 0xCE3E3E;
    public int Debug { get; set; } = 0xCE8800;
    public int Gray { get; set; } = 0x555555;
    public int SelfUserName { get; set; } = 0x2CACE8;
    public int Whisper { get; set; } = 0xAFAFAF;
    public int Highlight { get; set; } = 0x8D00CE;
    public int Red { get; set; } = 0xCE3E3E;
    public int Green { get; set; } = 0x00CE00;
    public int Cyan { get; set; } = 0x00FFFF;
    public int Speaker { get; set; } = 0x51CECE;
    public int Guest { get; set; } = 0x8D00CE;
    public int UserNameDefault { get; set; } = 0xA89D65;
    public int EmoteDefault { get; set; } = 0xA89D65;
}

/// <summary>A custom color scheme plus its display name, as exported to/imported from a standalone .json file to share with other users.</summary>
public sealed class NamedCustomPalette
{
    public string Name { get; set; } = "My Color Scheme";
    public CustomChatPalette Colors { get; set; } = new();
}

/// <summary>
/// Optional bridge relaying this bot's Battle.net chat to/from a Discord
/// channel. Not wired to any Discord client yet — this is the configuration
/// surface planned ahead of that work, per BotEngine.ChatMessage/Log already
/// being plain events any future subscriber (including a Discord relay) can
/// hook into without changes to the engine.
/// </summary>
public sealed class DiscordBridgeConfig
{
    public bool Enabled { get; set; }

    /// <summary>Discord bot token. Obfuscated at rest like Password — see ObfuscatedPasswordJsonConverter. Never logged or echoed back in chat.</summary>
    [JsonConverter(typeof(ObfuscatedPasswordJsonConverter))]
    public string BotToken { get; set; } = "";

    public ulong ChannelId { get; set; }

    /// <summary>Minimum delay between relayed messages in each direction, to avoid flooding either side.</summary>
    public int RelayDelaySeconds { get; set; } = 2;

    public bool RelayBattlenetToDiscord { get; set; } = true;

    public bool RelayDiscordToBattlenet { get; set; } = true;
}

