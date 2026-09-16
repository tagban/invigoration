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

    /// <summary>The modern tracker protocol. A v1/v2 tracker ignores the extra bytes a v3 client sends and answers with its own version, so asking costs nothing.</summary>
    public const ushort TrackerVersion3 = 0x0003;

    /// <summary>"H3" — marks a v3 extension block inside an otherwise v1-shaped registration datagram.</summary>
    public const ushort TrackerV3ExtensionMagic = 0x4833;

    /// <summary>What a v3 client asks for in its handshake; the tracker answers with its own set and the overlap is what's in play.</summary>
    [Flags]
    public enum TrackerFeatures : ushort
    {
        None = 0,

        /// <summary>Server records may carry IPv6 addresses.</summary>
        IPv6 = 0x0001,

        /// <summary>Search text and pagination can be sent with a listing request.</summary>
        Query = 0x0002,

        ClientAuth = 0x0004,
        RegistrationAck = 0x0008,
        Hmac = 0x0010,
    }

    /// <summary>TLV ids for the query parameters a v3 listing request can carry.</summary>
    public const ushort TrackerQuerySearchText = 0x1001;
    public const ushort TrackerQueryPageOffset = 0x1010;
    public const ushort TrackerQueryPageLimit = 0x1011;
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
    // --- Files. Listing is a plain transaction; the bytes themselves move over a second,
    // short-lived connection on port+1 (see HotlineFileTransfer). ---

    /// <summary>Client -&gt; server: list one directory. Reply carries a FileNameWithInfo field per entry.</summary>
    GetFileNameList = 200,

    /// <summary>Client -&gt; server: ask to download a file. The reply's TransferRefNum is what the transfer connection presents.</summary>
    DownloadFile = 202,

    /// <summary>Client -&gt; server: ask to upload a file. Same shape as DownloadFile — the reply hands back a reference number to push the bytes with.</summary>
    UploadFile = 203,

    /// <summary>Client -&gt; server: download a whole folder, recursively. The transfer connection then walks its contents item by item — see HotlineFolderTransfer.</summary>
    DownloadFolder = 210,

    /// <summary>Client -&gt; server: upload a whole folder.</summary>
    UploadFolder = 213,

    // --- News, threaded (Hotline 1.5+): categories and bundles form a tree, each category holding
    // articles. See HotlineNewsPath for how a location in that tree goes on the wire. ---

    /// <summary>Client -&gt; server: list the categories/bundles at one path.</summary>
    GetNewsCategoryNameList = 370,

    /// <summary>Client -&gt; server: list the articles in one category.</summary>
    GetNewsArticleNameList = 371,

    /// <summary>Client -&gt; server: fetch one article's body.</summary>
    GetNewsArticleData = 400,

    /// <summary>Client -&gt; server: post an article (or a reply to one).</summary>
    PostNewsArticle = 410,

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
    // --- Files ---

    /// <summary>Reply to GetFileNameList, repeated once per entry — see HotlineFileEntry for the packed layout.</summary>
    FileNameWithInfo = 200,

    FileName = 201,

    /// <summary>A directory, packed as a count followed by one length-prefixed component each — see HotlineNewsPath, which uses the same encoding.</summary>
    FilePath = 202,

    /// <summary>Total bytes the transfer will carry (the flattened file, forks and headers included) — not the data fork's own size.</summary>
    TransferSize = 108,

    FileSize = 207,

    /// <summary>The token the separate transfer connection presents to claim this transfer.</summary>
    TransferRefNum = 107,

    /// <summary>How many transfers are queued ahead of this one; 0 means it can start now. (116 — an earlier draft here had 212, which is File New Path.)</summary>
    WaitingCount = 116,

    /// <summary>How many items a folder transfer will carry, so both sides know when the item loop is done.</summary>
    FolderItemCount = 220,

    /// <summary>Flags for a transfer — resume and so on. Sent on an upload request.</summary>
    FileTransferOptions = 204,

    /// <summary>Where a partially-transferred file left off, for resuming.</summary>
    FileResumeData = 203,

    // --- News (threaded) ---

    /// <summary>Where in the news tree a request applies — same packed encoding as FilePath. Absent means the root.</summary>
    NewsPath = 325,

    /// <summary>Reply to GetNewsCategoryNameList: the packed category/bundle listing.</summary>
    NewsCategoryListData = 323,

    /// <summary>Reply to GetNewsArticleNameList: the packed article listing.</summary>
    NewsArticleListData = 321,

    NewsArticleId = 326,

    /// <summary>MIME type of an article body — servers use "text/plain" in practice.</summary>
    NewsArticleDataFlavor = 327,

    NewsArticleTitle = 328,
    NewsArticlePoster = 329,
    NewsArticleDate = 330,

    /// <summary>The article being replied to, when posting a reply rather than a new thread.</summary>
    NewsArticleParent = 335,

    /// <summary>The article body itself.</summary>
    NewsArticleData = 333,

    // --- HOPE (Hotline One-time Password Extension), the "3.x"-era secure login. See HotlineHope. ---

    /// <summary>DATA_HOPE_APP_ID — the client's 4-character application identifier.</summary>
    HopeAppId = 0x0E01,

    /// <summary>DATA_HOPE_APP_STRING — a human-readable name and version. HOPE is the only way a client can tell a server what it is.</summary>
    HopeAppString = 0x0E02,

    /// <summary>DATA_HOPE_SESSION_KEY — the server's 64-byte challenge, which the password is MAC'd against.</summary>
    HopeSessionKey = 0x0E03,

    /// <summary>DATA_HOPE_MAC_ALGORITHM — the client's list of supported MACs, and the server's single choice in the reply.</summary>
    HopeMacAlgorithm = 0x0E04,

    /// <summary>DATA_HOPE_SERVER_CIPHER — the cipher the server will encrypt with, when transport encryption is negotiated.</summary>
    HopeServerCipher = 0x0EC1,

    /// <summary>DATA_HOPE_CLIENT_CIPHER — the cipher the client will encrypt with.</summary>
    HopeClientCipher = 0x0EC2,

    HopeServerCipherMode = 0x0EC3,
    HopeClientCipherMode = 0x0EC4,
    HopeServerChecksum = 0x0EC7,
    HopeClientChecksum = 0x0EC8,

    /// <summary>DATA_HOPE_SERVER_COMPRESSION. Never sent as an empty list — per the spec that crashes some clients outright.</summary>
    HopeServerCompression = 0x0EC9,

    HopeClientCompression = 0x0ECA,

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
