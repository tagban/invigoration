namespace Invigoration.Core.Config;

/// <summary>Per-bot settings. Replaces the VB6 BotData/BotNetData globals.</summary>
public sealed class BotConfig
{
    public string DisplayName { get; set; } = "New Bot";
    public string Username { get; set; } = "";
    public string Password { get; set; } = "";
    public string CdKey { get; set; } = "";
    public string BotMaster { get; set; } = "";
    public string Trigger { get; set; } = "!";
    public bool UseUdp { get; set; }
    public string Email { get; set; } = "";
    public PingDisplayMode ShowPing { get; set; } = PingDisplayMode.Bars;
    public bool JoinNotify { get; set; }
    public string BattlenetServer { get; set; } = "";
    public int BattlenetPort { get; set; } = 6112;
    public string BnlsServer { get; set; } = "";
    public int BnlsPort { get; set; } = 9367;
    public string HomeChannel { get; set; } = "";

    /// <summary>4-character BNCS product code, e.g. "VD2D" = Diablo II, "PX2D" = Diablo II: LoD.</summary>
    public string Product { get; set; } = "VD2D";

    public string Realm { get; set; } = "";
    public bool ZeroPing { get; set; }
    public bool NegPing { get; set; }
    public bool ShowBnccIcon { get; set; }

    /// <summary>
    /// Full BNCS binary protocol (needs BNLS for hashing) vs. the plain-text
    /// chat-gateway/telnet protocol (username + plaintext password only, no
    /// BNLS/CD-key/version-check at all). Official Battle.net disabled the
    /// gateway protocol in 2005, so TelnetGateway only works against PVPGN.
    /// </summary>
    public ConnectionMode ConnectionMode { get; set; } = ConnectionMode.BncsBinary;

    public DiscordBridgeConfig Discord { get; set; } = new();
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

    /// <summary>Discord bot token. Secret — stored locally like the BNCS password; never logged or echoed back in chat.</summary>
    public string BotToken { get; set; } = "";

    public ulong ChannelId { get; set; }

    /// <summary>Minimum delay between relayed messages in each direction, to avoid flooding either side.</summary>
    public int RelayDelaySeconds { get; set; } = 2;

    public bool RelayBattlenetToDiscord { get; set; } = true;

    public bool RelayDiscordToBattlenet { get; set; } = true;
}

public enum PingDisplayMode
{
    Numeric,
    Bars,
    BarsCompact,
}

public enum ConnectionMode
{
    BncsBinary,
    TelnetGateway,
}
