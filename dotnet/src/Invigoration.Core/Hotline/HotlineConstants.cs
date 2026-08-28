namespace Invigoration.Core.Hotline;

/// <summary>
/// Opcodes/field IDs/ports for the classic Hotline protocol (HTLC over TCP, plus the separate HTRK
/// tracker protocol) — ported from the actively-maintained Hotline-Navigator client
/// (github.com/bourbonicfisky/Hotline-Navigator, a real Rust/Tauri client that speaks both the
/// legacy 1.x/1.5/1.8 wire format and the newer "Hotline 3.x" HOPE-encrypted variant), which the
/// user pointed at as the reference for "the latest spec supported... while still supporting
/// older formatting/structure." Only the legacy (unencrypted) subset is implemented here — see
/// HotlineTransactionClient's remarks for why HOPE is out of scope for now.
/// </summary>
public static class HotlineConstants
{
    public static readonly byte[] ProtocolId = "TRTP"u8.ToArray();
    public static readonly byte[] SubProtocolId = "HOTL"u8.ToArray();
    public const ushort ProtocolVersion = 0x0001;
    public const ushort ProtocolSubversion = 0x0002;

    public const int TransactionHeaderSize = 20;

    public const ushort DefaultServerPort = 5500;
    public const ushort DefaultTrackerPort = 5498;

    public static readonly byte[] TrackerMagic = "HTRK"u8.ToArray();
    public const ushort TrackerVersion = 0x0001;
}

/// <summary>The subset of Hotline's TransactionType opcodes this client actually sends/handles — chat, login, user list, and keepalive. File transfer, news, and media transactions are a known, accepted gap (see HotlineTransactionClient's remarks).</summary>
public enum HotlineTransactionType : ushort
{
    Reply = 0,
    Error = 100,
    ServerMessage = 104,
    SendChat = 105,
    ChatMessage = 106,
    Login = 107,
    ShowAgreement = 109,
    DisconnectUser = 110,
    DisconnectMessage = 111,
    NotifyChatOfUserChange = 117,
    NotifyChatOfUserDelete = 118,
    Agreed = 121,
    GetUserNameList = 300,
    NotifyUserChange = 301,
    NotifyUserDelete = 302,
    SetClientUserInfo = 304,
    UserAccess = 354,
    KeepAlive = 500,

    /// <summary>Client -&gt; server (request/reply): fetch a batch of persisted chat history. Only meaningful once the server has confirmed CAPABILITY_CHAT_HISTORY (see HotlineCapabilityBits) during login — a pre-2.5 (1.2.3+) server never understands this and this client never sends it unless the login reply actually echoed the bit back. See github.com/fogWraith/Hotline/blob/main/Docs/Protocol/Capabilities-Chat-History.md.</summary>
    GetChatHistory = 700,
}

/// <summary>The subset of Hotline's FieldType parameter IDs this client actually reads/writes.</summary>
public enum HotlineFieldType : ushort
{
    ErrorText = 100,
    Data = 101,
    UserName = 102,
    UserId = 103,
    UserIconId = 104,
    UserLogin = 105,
    UserPassword = 106,
    ChatOptions = 109,
    Options = 113,
    ChatId = 114,
    ServerAgreement = 150,
    NoServerAgreement = 154,
    VersionNumber = 160,
    ServerName = 162,
    UserNameWithInfo = 300,
    UserFlags = 112,
    NickColor = 1280,
    UserAccess = 110,

    /// <summary>DATA_CAPABILITIES (0x01F0) — bitmask a client advertises at Login and the server echoes (only the bits it actually confirms) in the reply. See HotlineCapabilityBits and github.com/fogWraith/Hotline/blob/main/Docs/Protocol/Capabilities.md.</summary>
    Capabilities = 0x01F0,

    /// <summary>DATA_CHANNEL_ID — Get Chat History (700)'s target channel; 0 is always the public chat, the only one this client (or most real servers yet) actually has.</summary>
    ChannelId = 0x0F01,

    /// <summary>DATA_HISTORY_BEFORE (uint64) — pagination cursor: messages with IDs strictly less than this.</summary>
    HistoryBefore = 0x0F02,

    /// <summary>DATA_HISTORY_AFTER (uint64) — pagination cursor: messages with IDs strictly greater than this.</summary>
    HistoryAfter = 0x0F03,

    /// <summary>DATA_HISTORY_LIMIT (uint16) — max messages to return in one Get Chat History reply.</summary>
    HistoryLimit = 0x0F04,

    /// <summary>DATA_HISTORY_ENTRY (binary, repeated 0-N per reply) — one packed HotlineChatHistoryEntry. See its own Parse remarks for the wire layout.</summary>
    HistoryEntry = 0x0F05,

    /// <summary>DATA_HISTORY_HAS_MORE (uint8) — 1 if more messages exist beyond this batch in the direction queried.</summary>
    HistoryHasMore = 0x0F06,

    /// <summary>DATA_HISTORY_MAX_MSGS (uint32, login reply only) — server's retention policy, informational; 0 = unlimited.</summary>
    HistoryMaxMsgs = 0x0F07,

    /// <summary>DATA_HISTORY_MAX_DAYS (uint32, login reply only) — server's retention policy, informational; 0 = unlimited.</summary>
    HistoryMaxDays = 0x0F08,
}

/// <summary>
/// Bits within DATA_CAPABILITIES (field 0x01F0) — the login-time feature-negotiation bitmask a
/// small but growing "2.5"-era community spec defines (github.com/fogWraith/Hotline/blob/main/Docs/Protocol/Capabilities.md).
/// Only the bit this client actually implements is named; the spec defines several others (large
/// files, UTF-8 text encoding, voice, inline media, extended privileges, messaging, modern dates)
/// this client doesn't speak yet. A server that doesn't recognize DATA_CAPABILITIES at all (any
/// pre-2.5, 1.2.3-and-up server) simply never echoes it back — SupportsChatHistory stays false and
/// nothing about this client's behavior changes for it.
/// </summary>
public static class HotlineCapabilityBits
{
    public const int ChatHistory = 0x0010;
}

/// <summary>
/// Bit positions within the 64-bit account-access bitmap the server sends us (only about
/// ourselves, via the UserAccess/354 transaction — never broadcast for other users; see
/// HotlineTransactionClient's remarks) — confirmed against hlwiki.com/index.php/AccessPriviledges,
/// the same standard bit layout Mobius's own source uses. Only the bits this client actually
/// checks are named; the rest of the 64 bits exist but aren't needed yet.
/// </summary>
public static class HotlineAccessBits
{
    public const int DeleteFile = 0;
    public const int UploadFile = 1;
    public const int DownloadFile = 2;
    public const int ReadChat = 9;
    public const int SendChat = 10;
    public const int CreateUser = 14;
    public const int DeleteUser = 15;
    public const int OpenUser = 16;
    public const int ModifyUser = 17;
    public const int NewsReadArticle = 20;
    public const int NewsPostArticle = 21;
    public const int DisconnectUser = 22;
    public const int CannotBeDisconnected = 23;
    public const int AnyName = 26;
    public const int NoAgreement = 27;
}

/// <summary>
/// Bit positions (not values — shift by these) within the 2-byte per-user UserFlags field, the
/// only per-*other*-user status info the protocol actually broadcasts (see HotlineAccessBits'
/// remarks for the much richer 64-bit access bitmap, which is self-only). Confirmed against
/// Mobius's real source (hotline/user.go): only Admin, Away, RefusePM, RefusePChat exist — no
/// separate "Mod" bit at the wire level.
/// </summary>
public static class HotlineUserFlagBits
{
    public const int Away = 0;
    public const int Admin = 1;
    public const int RefusePm = 2;
    public const int RefusePChat = 3;
}
