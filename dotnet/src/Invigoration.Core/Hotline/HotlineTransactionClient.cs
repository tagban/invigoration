using System.Buffers.Binary;
using System.Collections.Concurrent;
using System.Text;
using Invigoration.Core.Networking;

namespace Invigoration.Core.Hotline;

/// <summary>
/// A real Hotline (HTLC) client connection: TRTP/HOTL handshake, legacy (XOR-obfuscated, not
/// HOPE-encrypted) login, and chat send/receive — built directly on FramedTcpClient the same way
/// BncsConnection/BnlsConnection are, with TryGetFrameLength switching shape once (8 bytes for the
/// one-time handshake reply, then the real 20-byte-header transaction framing for everything
/// after — see HotlineTransactionFrame.TryGetFrameLength).
///
/// Deliberately scoped to what a chat bot actually needs: login, chat, and a live user list.
/// Out of scope for now (known, accepted gaps, not oversights):
/// - HOPE encryption (Hotline 3.x's optional secure-login extension) — Hotline-Navigator's own
///   docs describe it as actively harmful to attempt against a server that doesn't support it
///   ("poisons the connection... they treat it as a failed login"), so this client always uses
///   the legacy XOR login every other Hotline client still understands.
/// - File transfer, news, and media transactions.
/// </summary>
public sealed class HotlineTransactionClient : FramedTcpClient
{
    /// <summary>Real clients send one of these every ~180s to keep the connection from timing out — confirmed against Hotline-Navigator's source (it falls back to a GetUserNameList poll only for older servers that don't tolerate a bare KeepAlive, which this client doesn't need to special-case).</summary>
    private static readonly TimeSpan KeepAliveInterval = TimeSpan.FromSeconds(180);

    private readonly ConcurrentDictionary<uint, TaskCompletionSource<HotlineTransactionFrame>> _pendingReplies = new();
    private readonly List<HotlineUser> _users = [];

    private bool _handshakeComplete;
    private TaskCompletionSource<bool>? _handshakeTcs;
    private uint _nextTransactionId = 1;
    private CancellationTokenSource? _keepAliveCts;
    private TaskCompletionSource<bool>? _agreementArrivedTcs;
    private bool _userListDeferredForAgreement;
    private string _nickname = "";
    private ushort _iconId;

    public event Action<string>? ChatMessageReceived;
    public event Action<string>? ServerMessageReceived;

    /// <summary>Fired once, right after login, with the room's initial member list.</summary>
    public event Action<IReadOnlyList<HotlineUser>>? UserListReceived;

    /// <summary>Fired for both a genuinely new user joining and an existing user's name/icon/flags changing — same as the server doesn't distinguish the two cases at the wire level either (NotifyUserChange covers both).</summary>
    public event Action<HotlineUser>? UserChanged;

    public event Action<ushort>? UserLeft;

    /// <summary>A server prompted its agreement/rules text and AutoAcceptAgreement is off — the UI is expected to show it and call AcceptAgreementAsync only on an explicit user action. Never fired when AutoAcceptAgreement is on (that path sends Agreed immediately instead).</summary>
    public event Action<string>? AgreementReceived;

    /// <summary>One inbound transaction failed to parse — the connection stays alive (see OnPacketReceived's remarks); this is purely informational for surfacing/logging.</summary>
    public event Action<Exception>? ProtocolError;

    /// <summary>A server sent DisconnectMessage right before closing the connection — the real protocol's way of explaining why (kicked, banned, duplicate login, etc.). Surfaced separately from ServerMessageReceived since it's specifically the last thing said before a disconnect, not routine chatter.</summary>
    public event Action<string>? DisconnectMessageReceived;

    /// <summary>Every inbound transaction, decoded — type name/number and every field — fired only when Debug is on. Exists specifically to diagnose "a server disconnects us and we don't know why": the last few lines before Disconnected fires are whatever the server actually sent right before closing the connection.</summary>
    public event Action<string>? DebugLog;

    /// <summary>
    /// Why the last <see cref="ConnectAndLoginAsync"/> was turned down by the server, or null when
    /// it succeeded or never got that far (unreachable host, failed handshake). The distinction is
    /// what lets auto-reconnect keep retrying a server that's merely down while standing down for
    /// one that refused the credentials.
    /// </summary>
    public string? LoginRefusedReason { get; private set; }

    public IReadOnlyList<HotlineUser> Users => _users;

    /// <summary>The server's own reported name (login reply field 162, e.g. "MacDomain", "Hotline Central Hub" — confirmed live), null until login succeeds. A saved profile's or tracker listing's own name should still win in the UI when one exists; this is the fallback for a session that has neither.</summary>
    public string? ServerName { get; private set; }

    /// <summary>Our own 64-bit account-access bitmap, from the unsolicited UserAccess(354) transaction the server sends right after connecting — see HotlineAccessBits' remarks on why this is our own permissions only, never another user's.</summary>
    public ulong OwnAccessBits { get; private set; }

    /// <summary>True once login completes if the server confirmed CAPABILITY_CHAT_HISTORY (DATA_CAPABILITIES bit 4) — only then is it meaningful to call GetChatHistoryAsync. False for any pre-2.5 (1.2.3+) server, which never echoes DATA_CAPABILITIES at all.</summary>
    public bool SupportsChatHistory { get; private set; }

    /// <summary>The server's advertised retention policy from the login reply (DATA_HISTORY_MAX_MSGS/MAX_DAYS) — informational only, per the spec's own remarks; null if the server didn't include it (or chat history isn't supported at all).</summary>
    public uint? HistoryMaxMessages { get; private set; }

    public uint? HistoryMaxDays { get; private set; }

    /// <summary>
    /// Whether one privilege bit is set in the account-access bitmap the server sent at login.
    ///
    /// The bits are numbered from the MOST significant bit of the FIRST byte — bit 2 is byte 0's
    /// 0x20, not 0x04 of the last byte — matching Mobius's own <c>bits[i/8] &amp; (1&lt;&lt;(7-i%8))</c>.
    /// The protocol docs don't state the convention at all, so it was settled against a real
    /// server: MacDomain hands a guest 60700c2003800000, which read this way is exactly
    /// upload/download, read/send chat, read/post news and any-name — and read the other way round
    /// claims a guest can't be disconnected while denying the downloads it plainly allows.
    ///
    /// Reading it the wrong way is what made the Files tab report "this account isn't allowed to
    /// browse files" to a guest who could (2026-09-16).
    /// </summary>
    public bool HasOwnAccess(int bit) =>
        bit is >= 0 and < 64 && (OwnAccessBits & (1UL << (63 - bit))) != 0;

    /// <summary>Off by default — never silently agree to a server's rules on the user's behalf. Set before ConnectAndLoginAsync; per-tracker, from HotlineTrackerConfig.AutoAcceptAgreement.</summary>
    public bool AutoAcceptAgreement { get; set; }

    /// <summary>
    /// Whether to offer HOPE's secure login (see <see cref="HotlineHope"/>) before falling back to
    /// the classic bitwise-inverted password. Off by default and opt-in per server: it costs an
    /// extra round trip against a server that doesn't speak it, and an unfamiliar login shape is a
    /// stability risk against the range of real servers out there.
    /// </summary>
    public bool UseSecureLogin { get; set; }

    /// <summary>Which MAC a HOPE login actually used, or null when the login was the classic one — shown in the session log so it's clear whether the password crossed the wire obfuscated or MAC'd.</summary>
    public string? SecureLoginAlgorithm { get; private set; }

    /// <summary>Where this client actually connected, for comparing against the address a HOPE session key claims.</summary>
    private (System.Net.IPAddress Address, int Port)? _connectedEndpoint;

    /// <summary>The server's self-reported name and version, which only HOPE exposes. Null on a classic login.</summary>
    public string? ServerApplication { get; private set; }

    /// <summary>Raised when the address embedded in a HOPE session key isn't the one we connected to — a NAT or proxy in the path, possibly something worse. Advisory: the login still proceeds.</summary>
    public event Action<string>? SecureLoginAddressMismatch;

    /// <summary>Off by default — logs every inbound transaction via DebugLog. Set before ConnectAndLoginAsync; per-tracker, from HotlineTrackerConfig.Debug.</summary>
    public bool Debug { get; set; }

    public HotlineTransactionClient()
    {
        PacketReceived += OnPacketReceived;
        Disconnected += _ =>
        {
            _keepAliveCts?.Cancel();
            // A server that rejects our connection outright — no bytes at all, just an instant
            // reset — never reaches OnPacketReceived, so without this the handshake attempt would
            // sit on its full 10s timeout instead of failing (and retrying with the legacy
            // subversion) right away.
            _handshakeTcs?.TrySetResult(false);
        };
    }

    /// <summary>
    /// Connects, performs the TRTP/HOTL handshake, then logs in with a legacy XOR-obfuscated
    /// login/password (empty password is sent as no UserPassword field at all, same as a real
    /// client asked to log in anonymously). Returns false on a handshake failure, a login error
    /// (bad credentials, banned, server full — the specific HotlineTransactionFrame.ErrorCode from
    /// the reply is swallowed here since a chat bot just needs yes/no; a future UI can surface it).
    /// </summary>
    public async Task<bool> ConnectAndLoginAsync(string host, int port, string login, string password, string nickname, ushort iconId, ushort? clientVersion = 6112, bool advertiseChatHistory = false, CancellationToken ct = default)
    {
        _nickname = nickname;
        _iconId = iconId;
        LoginRefusedReason = null;

        // Created before anything else, not after the login reply — confirmed live that a real
        // server can (and does) push ShowAgreement before its own Login reply arrives, and this
        // needs to exist in time to catch that or the signal is silently dropped (a real bug this
        // exact ordering caused: the null-conditional TrySetResult on a not-yet-created TCS is a
        // no-op, so the later grace-window wait just timed out believing no agreement existed).
        _agreementArrivedTcs = new TaskCompletionSource<bool>();

        // Try the modern subversion first, then a fresh reconnect with the legacy one if the
        // server rejects it — confirmed against Hotline-Navigator's real establish_connection():
        // a server that doesn't understand subversion 2 sends back a non-zero error code in the
        // 8-byte handshake reply (not a silent disconnect), and the fix is a brand-new TCP
        // connection, not resending the handshake on the same socket. This is almost certainly
        // why an older Mobius-based server disconnected instantly against the old
        // always-subversion-2 code — same class of "1.2.3-modern server structure" compatibility
        // gap the user flagged directly.
        if (!await TryHandshakeAsync(host, port, HotlineConstants.ProtocolSubversion, ct).ConfigureAwait(false))
        {
            if (!await TryHandshakeAsync(host, port, (ushort)0x0001, ct).ConfigureAwait(false))
            {
                return false;
            }
        }

        // Ask whether this server speaks HOPE before sending anything secret. A server that does
        // answers with a challenge to MAC the password against; one that doesn't just refuses the
        // null login, costing one round trip and telling it nothing. Opt-in per server
        // (HotlineServerProfile.UseSecureLogin) for the same reason the chat-history capability is:
        // an unfamiliar login shape against the wide range of real servers out there is a
        // connection-stability risk, not a free win.
        if (UseSecureLogin &&
            await TrySecureLoginAsync(login, password, nickname, iconId, clientVersion, advertiseChatHistory, ct).ConfigureAwait(false) is { } secureResult)
        {
            return secureResult;
        }

        List<HotlineField> loginFields =
        [
            new HotlineField(HotlineFieldType.UserLogin, XorObfuscate(login)),
        ];
        if (!string.IsNullOrEmpty(password))
        {
            loginFields.Add(new HotlineField(HotlineFieldType.UserPassword, XorObfuscate(password)));
        }

        loginFields.Add(new HotlineField(HotlineFieldType.UserIconId, iconId));
        loginFields.Add(new HotlineField(HotlineFieldType.UserName, nickname));
        // 6112 by default — deliberately not a real Hotline release's version number (contrast the
        // old 150/1.5.x-honesty reasoning this replaces). Per explicit request: VersionNumber
        // reveals a lot about the connecting client to a modern server (see the protocol docs at
        // github.com/fogWraith/Hotline/tree/main/Docs/Protocol), so this is chosen specifically to
        // be a distinctive, unused-by-any-real-client number that uniquely identifies Invigoration
        // itself. Overridable per-server (HotlineServerProfile.ClientVersion) for testing how
        // different real servers react to different claimed versions. Null omits the field
        // entirely — also per explicit request, for a real newer server build that expects no
        // VersionNumber field at all (not just a specific value).
        if (clientVersion.HasValue)
        {
            loginFields.Add(new HotlineField(HotlineFieldType.VersionNumber, clientVersion.Value));
        }

        // Off by default, per-server opt-in (HotlineServerProfile.AdvertiseChatHistorySupport) —
        // TLV framing means an unrecognized field SHOULD be safely skippable by any pre-2.5
        // server, but this field is new/unproven against the wide range of real server
        // implementations out there, and intermittent forced disconnects started appearing right
        // around when this was first added. Not worth risking connection stability for a
        // nice-to-have feature nobody explicitly asked to always have on.
        if (advertiseChatHistory)
        {
            loginFields.Add(new HotlineField(HotlineFieldType.Capabilities, (ushort)HotlineCapabilityBits.ChatHistory));
        }

        var loginReply = await SendTransactionAsync(HotlineTransactionType.Login, [.. loginFields], ct).ConfigureAwait(false);
        if (loginReply is not { ErrorCode: 0 })
        {
            // The server answered and said no — bad credentials, banned, or full. Worth
            // distinguishing from "couldn't reach it": reconnecting can't talk a refusal round,
            // and a run of rejected logins is what gets an account or address banned.
            LoginRefusedReason = loginReply?.Field(HotlineFieldType.ErrorText)?.AsString() is { Length: > 0 } text
                ? text
                : "the server refused the login";
            return false;
        }

        return await FinishLoginAsync(loginReply, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Everything that happens once a login — classic or HOPE — has been accepted: read what the
    /// server reported about itself, settle any agreement it pushed, fetch the user list, and
    /// start the keepalive. Shared so the two login paths can't drift apart.
    /// </summary>
    private async Task<bool> FinishLoginAsync(HotlineTransactionFrame loginReply, CancellationToken ct)
    {
        var serverName = loginReply.Field(HotlineFieldType.ServerName)?.AsString();
        if (!string.IsNullOrEmpty(serverName))
        {
            ServerName = serverName;
        }

        // The server only echoes back the bits it actually confirms — absent entirely means
        // "standard mode," per the spec's own absence-handling rule, not "everything denied but
        // present as zero." Either way SupportsChatHistory correctly ends up false.
        var confirmedCapabilities = loginReply.Field(HotlineFieldType.Capabilities)?.AsUInt16() ?? 0;
        SupportsChatHistory = (confirmedCapabilities & HotlineCapabilityBits.ChatHistory) != 0;
        HistoryMaxMessages = loginReply.Field(HotlineFieldType.HistoryMaxMsgs)?.AsUInt32();
        HistoryMaxDays = loginReply.Field(HotlineFieldType.HistoryMaxDays)?.AsUInt32();

        // Confirmed live: a real Mobius-based server disconnected a session that requested
        // GetUserNameList without first resolving an agreement it had just pushed — plausibly its
        // own anti-bot heuristic ("a real client wouldn't query the room before agreeing to its
        // rules"). Give a server-pushed ShowAgreement a brief window to arrive (empirically it
        // arrives near-instantly, right alongside the login reply, if it's coming at all) before
        // deciding whether it's safe to fetch the user list now. (_agreementArrivedTcs was
        // created at the very top of this method, not here — see its remarks.)
        using (var graceCts = CancellationTokenSource.CreateLinkedTokenSource(ct))
        {
            graceCts.CancelAfter(TimeSpan.FromMilliseconds(800));
            try
            {
                await _agreementArrivedTcs.Task.WaitAsync(graceCts.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (!ct.IsCancellationRequested)
            {
                // No agreement showed up in the grace window — nothing to wait for.
            }
        }

        if (_agreementArrivedTcs.Task is { IsCompletedSuccessfully: true })
        {
            // Defer first — AutoAcceptAgreement's own AcceptAgreementAsync call below fetches the
            // user list itself once Agreed is actually sent, in the correct order. Without
            // AutoAcceptAgreement, it stays deferred until the user explicitly accepts.
            _userListDeferredForAgreement = true;
            if (AutoAcceptAgreement)
            {
                await AcceptAgreementAsync(ct).ConfigureAwait(false);
            }
        }
        else
        {
            await FetchUserListAsync(ct).ConfigureAwait(false);
        }

        // The connection can die at any point during the sequence above (login reply, the
        // agreement grace-window wait, an auto-accept send) without any single await throwing —
        // FramedTcpClient's Disconnected event just fires independently on the receive loop.
        // Confirmed live as a real bug: without this check, a session that died mid-login still
        // got reported "Connected." (and started its keepalive loop) purely because nothing it
        // awaited happened to throw.
        if (!IsConnected)
        {
            return false;
        }

        _keepAliveCts = new CancellationTokenSource();
        _ = KeepAliveLoopAsync(_keepAliveCts.Token);

        return true;
    }

    private async Task FetchUserListAsync(CancellationToken ct)
    {
        var userListReply = await SendTransactionAsync(HotlineTransactionType.GetUserNameList, [], ct).ConfigureAwait(false);
        if (userListReply is not null)
        {
            _users.Clear();
            _users.AddRange(userListReply.Fields.Where(f => f.Type == (ushort)HotlineFieldType.UserNameWithInfo).Select(f => HotlineUser.Parse(f.Data)));
            UserListReceived?.Invoke(_users);
        }
    }

    /// <summary>Opens a fresh connection and attempts the TRTP/HOTL handshake with the given subversion, returning whether the server accepted it (a non-zero error code in its 8-byte reply, or a timeout, both count as rejected).</summary>
    private async Task<bool> TryHandshakeAsync(string host, int port, ushort subversion, CancellationToken ct)
    {
        _handshakeComplete = false;
        _handshakeTcs = new TaskCompletionSource<bool>();

        await ConnectAsync(host, port, ct).ConfigureAwait(false);
        _connectedEndpoint = System.Net.IPAddress.TryParse(host, out var parsed) ? (parsed, port) : null;

        var handshake = new byte[12];
        HotlineConstants.ProtocolId.CopyTo(handshake, 0);
        HotlineConstants.SubProtocolId.CopyTo(handshake, 4);
        BinaryPrimitives.WriteUInt16BigEndian(handshake.AsSpan(8), HotlineConstants.ProtocolVersion);
        BinaryPrimitives.WriteUInt16BigEndian(handshake.AsSpan(10), subversion);
        await SendAsync(handshake, ct).ConfigureAwait(false);

        try
        {
            return await _handshakeTcs.Task.WaitAsync(TimeSpan.FromSeconds(10), ct).ConfigureAwait(false);
        }
        catch (TimeoutException)
        {
            return false;
        }
    }

    private async Task KeepAliveLoopAsync(CancellationToken ct)
    {
        try
        {
            while (!ct.IsCancellationRequested)
            {
                await Task.Delay(KeepAliveInterval, ct).ConfigureAwait(false);
                await SendAsync(HotlineTransactionFrame.Create(HotlineTransactionType.KeepAlive, NextId()).Encode(), ct).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException)
        {
            // Connection closed — nothing left to keep alive.
        }
    }

    public Task SendChatAsync(string message, CancellationToken ct = default) =>
        SendAsync(HotlineTransactionFrame.Create(HotlineTransactionType.SendChat, NextId(), new HotlineField(HotlineFieldType.Data, message)).Encode(), ct);

    /// <summary>
    /// Changes this client's own displayed nickname (and optionally icon) mid-session, via the
    /// real Hotline SetClientUserInfo(304) transaction — sendable any time after login, not just
    /// once at connect. The server responds by broadcasting NotifyUserChange(301) to the whole
    /// room (including us), which updates the Users list the normal way via UserChanged; this also
    /// updates the locally-cached _nickname/_iconId immediately so a later AcceptAgreementAsync
    /// resend uses the new values instead of the ones from login.
    /// </summary>
    public async Task ChangeUserInfoAsync(string nickname, ushort? iconId = null, CancellationToken ct = default)
    {
        _nickname = nickname;
        if (iconId.HasValue)
        {
            _iconId = iconId.Value;
        }

        await SendAsync(
            HotlineTransactionFrame.Create(
                HotlineTransactionType.SetClientUserInfo,
                NextId(),
                new HotlineField(HotlineFieldType.UserName, _nickname),
                new HotlineField(HotlineFieldType.UserIconId, _iconId)).Encode(),
            ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Explicitly agrees to a server's rules — either sent automatically (AutoAcceptAgreement) or
    /// in response to a real user action after AgreementReceived, never silently. Also fetches
    /// the user list if it was deferred waiting for exactly this (see ConnectAndLoginAsync's
    /// remarks). Resends UserName/UserIconID/Options — a real client's Agreed transaction isn't
    /// bare; Mobius's own HandleTranAgreed (confirmed against its actual source) reads these same
    /// three fields off this specific transaction, not just the earlier Login one. Options=0 (no
    /// bits set: not refusing PMs, not refusing chat, no auto-response) since this client doesn't
    /// support any of those yet.
    /// </summary>
    public async Task AcceptAgreementAsync(CancellationToken ct = default)
    {
        var agreed = HotlineTransactionFrame.Create(
            HotlineTransactionType.Agreed,
            NextId(),
            new HotlineField(HotlineFieldType.UserName, _nickname),
            new HotlineField(HotlineFieldType.UserIconId, _iconId),
            new HotlineField(HotlineFieldType.Options, (ushort)0));
        await SendAsync(agreed.Encode(), ct).ConfigureAwait(false);
        if (_userListDeferredForAgreement)
        {
            _userListDeferredForAgreement = false;
            await FetchUserListAsync(ct).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Fetches a batch of persisted chat history via Get Chat History (700) — only meaningful once
    /// SupportsChatHistory is true. No cursors (before/after both null) returns the most recent
    /// messages, oldest-first, exactly what's needed to pre-populate a session's chat log on
    /// connect instead of starting on a blank screen. Returns an empty, HasMore=false result
    /// (rather than throwing) on any error reply — a server that denies the request for
    /// permissions/config reasons shouldn't crash the connect flow, just silently skip history.
    /// </summary>
    public async Task<(IReadOnlyList<HotlineChatHistoryEntry> Entries, bool HasMore)> GetChatHistoryAsync(
        uint channelId = 0, ulong? before = null, ulong? after = null, ushort limit = 20, CancellationToken ct = default)
    {
        List<HotlineField> fields = [new HotlineField(HotlineFieldType.ChannelId, channelId)];
        if (before.HasValue)
        {
            fields.Add(new HotlineField(HotlineFieldType.HistoryBefore, before.Value));
        }

        if (after.HasValue)
        {
            fields.Add(new HotlineField(HotlineFieldType.HistoryAfter, after.Value));
        }

        fields.Add(new HotlineField(HotlineFieldType.HistoryLimit, limit));

        var reply = await SendTransactionAsync(HotlineTransactionType.GetChatHistory, [.. fields], ct).ConfigureAwait(false);
        if (reply is not { ErrorCode: 0 })
        {
            return ([], false);
        }

        var entries = reply.Fields
            .Where(f => f.Type == (ushort)HotlineFieldType.HistoryEntry)
            .Select(f => HotlineChatHistoryEntry.TryParse(f.Data))
            .Where(e => e is not null)
            .Select(e => e!)
            .ToList();

        var hasMore = reply.Field(HotlineFieldType.HistoryHasMore)?.AsBool() ?? false;
        return (entries, hasMore);
    }

    // --- Files. The transaction connection negotiates; the bytes move over their own connection
    // (see HotlineFileTransfer), which is why these return a reference number rather than data. ---

    /// <summary>Lists one directory — empty path for the server's root. Empty when the server says no (usually the account lacking the download privilege).</summary>
    public async Task<IReadOnlyList<HotlineFileEntry>> GetFileListAsync(
        IReadOnlyList<string>? path = null,
        CancellationToken ct = default)
    {
        var fields = new List<HotlineField>();
        if (HotlineNewsPath.ToFieldOrNull(HotlineFieldType.FilePath, path ?? []) is { } pathField)
        {
            fields.Add(pathField);
        }

        var reply = await SendTransactionAsync(HotlineTransactionType.GetFileNameList, [.. fields], ct).ConfigureAwait(false);
        if (reply is not { ErrorCode: 0 })
        {
            return [];
        }

        return reply.Fields
            .Where(f => f.Type == (ushort)HotlineFieldType.FileNameWithInfo)
            .Select(f => HotlineFileEntry.TryParse(f.Data))
            .Where(e => e is not null)
            .Select(e => e!)
            .ToList();
    }

    /// <summary>What the server hands back when it agrees to a transfer: the ticket, how much will move, and how many transfers are queued ahead.</summary>
    public sealed record TransferTicket(uint ReferenceNumber, uint TransferSize, uint FileSize, uint WaitingCount);

    /// <summary>
    /// Asks to download a file. Null when the server refuses (no such file, or no permission).
    /// The returned ticket is presented by <see cref="HotlineFileTransfer.DownloadAsync"/> on its
    /// own connection; a non-zero WaitingCount means the server has queued it behind others.
    /// </summary>
    public async Task<TransferTicket?> RequestDownloadAsync(
        IReadOnlyList<string> directory,
        string fileName,
        CancellationToken ct = default)
    {
        var fields = new List<HotlineField> { new(HotlineFieldType.FileName, fileName) };
        if (HotlineNewsPath.ToFieldOrNull(HotlineFieldType.FilePath, directory) is { } pathField)
        {
            fields.Add(pathField);
        }

        var reply = await SendTransactionAsync(HotlineTransactionType.DownloadFile, [.. fields], ct).ConfigureAwait(false);
        if (reply is not { ErrorCode: 0 } || reply.Field(HotlineFieldType.TransferRefNum) is not { } refNum)
        {
            return null;
        }

        return new TransferTicket(
            refNum.AsUInt32(),
            reply.Field(HotlineFieldType.TransferSize)?.AsUInt32() ?? 0,
            reply.Field(HotlineFieldType.FileSize)?.AsUInt32() ?? 0,
            reply.Field(HotlineFieldType.WaitingCount)?.AsUInt32() ?? 0);
    }

    /// <summary>Asks to upload a file into a directory. Null when the server refuses — usually the account lacking the upload privilege, or the folder being read-only.</summary>
    public async Task<TransferTicket?> RequestUploadAsync(
        IReadOnlyList<string> directory,
        string fileName,
        long fileSize,
        CancellationToken ct = default)
    {
        var fields = new List<HotlineField>
        {
            new(HotlineFieldType.FileName, fileName),
            new(HotlineFieldType.TransferSize, (uint)fileSize),
        };

        if (HotlineNewsPath.ToFieldOrNull(HotlineFieldType.FilePath, directory) is { } pathField)
        {
            fields.Insert(1, pathField);
        }

        var reply = await SendTransactionAsync(HotlineTransactionType.UploadFile, [.. fields], ct).ConfigureAwait(false);
        if (reply is not { ErrorCode: 0 } || reply.Field(HotlineFieldType.TransferRefNum) is not { } refNum)
        {
            return null;
        }

        return new TransferTicket(refNum.AsUInt32(), (uint)fileSize, (uint)fileSize, reply.Field(HotlineFieldType.WaitingCount)?.AsUInt32() ?? 0);
    }

    /// <summary>What a folder download needs: the ticket, plus how many items the server will walk through.</summary>
    public sealed record FolderTicket(uint ReferenceNumber, int ItemCount, uint TransferSize, uint WaitingCount);

    /// <summary>
    /// Asks to download a whole folder. Null when the server refuses. The item count matters —
    /// the folder transfer walks exactly that many items and the server sends no end marker.
    /// </summary>
    public async Task<FolderTicket?> RequestFolderDownloadAsync(
        IReadOnlyList<string> directory,
        string folderName,
        CancellationToken ct = default)
    {
        var fields = new List<HotlineField> { new(HotlineFieldType.FileName, folderName) };
        if (HotlineNewsPath.ToFieldOrNull(HotlineFieldType.FilePath, directory) is { } pathField)
        {
            fields.Add(pathField);
        }

        var reply = await SendTransactionAsync(HotlineTransactionType.DownloadFolder, [.. fields], ct).ConfigureAwait(false);
        if (reply is not { ErrorCode: 0 } || reply.Field(HotlineFieldType.TransferRefNum) is not { } refNum)
        {
            return null;
        }

        return new FolderTicket(
            refNum.AsUInt32(),
            (int)(reply.Field(HotlineFieldType.FolderItemCount)?.AsUInt32() ?? 0),
            reply.Field(HotlineFieldType.TransferSize)?.AsUInt32() ?? 0,
            reply.Field(HotlineFieldType.WaitingCount)?.AsUInt32() ?? 0);
    }

    /// <summary>Whether this account may download files, per the access bitmap from login.</summary>
    public bool CanDownloadFiles => HasOwnAccess(HotlineAccessBits.DownloadFile);

    /// <summary>Whether this account may upload files.</summary>
    public bool CanUploadFiles => HasOwnAccess(HotlineAccessBits.UploadFile);

    // --- News. Threaded news (Hotline 1.5+): bundles hold categories, categories hold articles,
    // and every request names where in that tree it applies (see HotlineNewsPath). A server
    // without news, or an account without the news privilege, answers with an error code, which
    // surfaces here as an empty list rather than an exception — asking is not a failure. ---

    /// <summary>
    /// The whole flat news document, or null if the server wouldn't give it. This is the original
    /// (1.x) news: one text blob, newest post first, rather than a tree of articles. Worth asking
    /// for even on a server that also has a threaded tree — on the classic servers still running,
    /// the real news is usually here and the tree is empty (MacDomain: 45 KB of posts here, two
    /// near-empty categories there).
    /// </summary>
    public async Task<string?> GetFlatNewsAsync(CancellationToken ct = default)
    {
        var reply = await SendTransactionAsync(HotlineTransactionType.GetMessages, [], ct).ConfigureAwait(false);
        return reply is { ErrorCode: 0 } ? reply.Field(HotlineFieldType.Data)?.AsString() ?? "" : null;
    }

    /// <summary>Adds a post to the flat news. Returns whether the server took it — a "no" is usually the account lacking the post privilege.</summary>
    public async Task<bool> PostFlatNewsAsync(string text, CancellationToken ct = default)
    {
        var reply = await SendTransactionAsync(
            HotlineTransactionType.PostFlatNews,
            [new HotlineField(HotlineFieldType.Data, text)],
            ct).ConfigureAwait(false);
        return reply is { ErrorCode: 0 };
    }

    /// <summary>A private message from another user, relayed by the server.</summary>
    public event Action<HotlinePrivateMessage>? PrivateMessageReceived;

    /// <summary>
    /// Sends a private message to one user. Hotline has no separate "whisper" the way Battle.net
    /// does — it's its own transaction addressed by the recipient's session id, which means a PM
    /// can only be sent to someone currently connected (ids are per-session, not per-account).
    /// </summary>
    public Task SendPrivateMessageAsync(ushort userId, string text, CancellationToken ct = default) =>
        SendTransactionAsync(
            HotlineTransactionType.SendInstantMessage,
            [
                new HotlineField(HotlineFieldType.UserId, userId),
                new HotlineField(HotlineFieldType.Options, (ushort)1),
                new HotlineField(HotlineFieldType.Data, text),
            ],
            ct);

    /// <summary>A post someone else just made, pushed by the server — just the new item, to go on top of what's already shown.</summary>
    public event Action<string>? FlatNewsPosted;

    /// <summary>The bundles and categories at <paramref name="path"/> — empty path for the root of the news tree.</summary>

    public async Task<IReadOnlyList<HotlineNewsCategory>> GetNewsCategoriesAsync(
        IReadOnlyList<string>? path = null,
        CancellationToken ct = default)
    {
        var fields = new List<HotlineField>();
        if (HotlineNewsPath.ToFieldOrNull(HotlineFieldType.NewsPath, path ?? []) is { } pathField)
        {
            fields.Add(pathField);
        }

        var reply = await SendTransactionAsync(HotlineTransactionType.GetNewsCategoryNameList, [.. fields], ct).ConfigureAwait(false);
        if (reply is not { ErrorCode: 0 })
        {
            return [];
        }

        // One field PER entry, not one field holding every entry — confirmed against MacDomain,
        // which sends a 10-byte field for its bundle and a 38-byte field for its category. Reading
        // only the first (what Field() returns) showed one of the two and silently dropped the rest.
        return reply.Fields
            .Where(f => f.Type == (ushort)HotlineFieldType.NewsCategoryListData)
            .SelectMany(f => HotlineNewsCategory.ParseList(f.Data))
            .ToList();
    }

    /// <summary>The articles in one category. Ordered as the server sent them — oldest first on every server seen so far.</summary>
    public async Task<IReadOnlyList<HotlineNewsArticle>> GetNewsArticlesAsync(
        IReadOnlyList<string> categoryPath,
        CancellationToken ct = default)
    {
        var fields = new List<HotlineField>();
        if (HotlineNewsPath.ToFieldOrNull(HotlineFieldType.NewsPath, categoryPath) is { } pathField)
        {
            fields.Add(pathField);
        }

        var reply = await SendTransactionAsync(HotlineTransactionType.GetNewsArticleNameList, [.. fields], ct).ConfigureAwait(false);
        if (reply is not { ErrorCode: 0 })
        {
            return [];
        }

        // Same again: a server may split the listing across several fields rather than one.
        return reply.Fields
            .Where(f => f.Type == (ushort)HotlineFieldType.NewsArticleListData)
            .SelectMany(f => HotlineNewsArticle.ParseList(f.Data))
            .ToList();
    }

    /// <summary>One article's body, or null if the server wouldn't give it up.</summary>
    public async Task<string?> GetNewsArticleBodyAsync(
        IReadOnlyList<string> categoryPath,
        uint articleId,
        string flavor = "text/plain",
        CancellationToken ct = default)
    {
        var fields = new List<HotlineField>
        {
            new(HotlineFieldType.NewsArticleId, articleId),
            new(HotlineFieldType.NewsArticleDataFlavor, string.IsNullOrEmpty(flavor) ? "text/plain" : flavor),
        };

        if (HotlineNewsPath.ToFieldOrNull(HotlineFieldType.NewsPath, categoryPath) is { } pathField)
        {
            fields.Insert(0, pathField);
        }

        var reply = await SendTransactionAsync(HotlineTransactionType.GetNewsArticleData, [.. fields], ct).ConfigureAwait(false);
        return reply is { ErrorCode: 0 } ? reply.Field(HotlineFieldType.NewsArticleData)?.AsString() ?? "" : null;
    }

    /// <summary>
    /// Posts an article, or a reply to one when <paramref name="parentArticleId"/> is given. Returns
    /// whether the server accepted it — a "no" is usually the account lacking the post privilege.
    /// </summary>
    public async Task<bool> PostNewsArticleAsync(
        IReadOnlyList<string> categoryPath,
        string title,
        string body,
        uint parentArticleId = 0,
        CancellationToken ct = default)
    {
        var fields = new List<HotlineField>
        {
            new(HotlineFieldType.NewsArticleTitle, title),
            new(HotlineFieldType.NewsArticleDataFlavor, "text/plain"),
            new(HotlineFieldType.NewsArticleData, body),
        };

        if (parentArticleId != 0)
        {
            fields.Add(new HotlineField(HotlineFieldType.NewsArticleParent, parentArticleId));
        }

        if (HotlineNewsPath.ToFieldOrNull(HotlineFieldType.NewsPath, categoryPath) is { } pathField)
        {
            fields.Insert(0, pathField);
        }

        var reply = await SendTransactionAsync(HotlineTransactionType.PostNewsArticle, [.. fields], ct).ConfigureAwait(false);
        return reply is { ErrorCode: 0 };
    }

    /// <summary>Whether this account may read news, per the access bitmap the server sent at login.</summary>
    public bool CanReadNews => HasOwnAccess(HotlineAccessBits.NewsReadArticle);

    /// <summary>Whether this account may post news.</summary>
    public bool CanPostNews => HasOwnAccess(HotlineAccessBits.NewsPostArticle);

    /// <summary>
    /// The HOPE exchange (see <see cref="HotlineHope"/>): announce ourselves with a null login,
    /// and if the server answers with a challenge, log in with the password MAC'd against it.
    ///
    /// Returns null — meaning "carry on with the classic login" — whenever this server turns out
    /// not to speak HOPE, or picked a MAC this client can't compute. Returns true/false only once
    /// a real HOPE login has been attempted and answered, so a refusal here is a genuine refusal
    /// and not something to retry the old way (which would send the password obfuscated to a
    /// server that just proved it wanted a MAC).
    /// </summary>
    private async Task<bool?> TrySecureLoginAsync(
        string login,
        string password,
        string nickname,
        ushort iconId,
        ushort? clientVersion,
        bool advertiseChatHistory,
        CancellationToken ct)
    {
        var hello = await SendTransactionAsync(HotlineTransactionType.Login, HotlineHope.IdentificationFields(), ct).ConfigureAwait(false);
        if (HotlineHope.ReadServerIdentification(hello) is not { } identification)
        {
            DebugLog?.Invoke("Secure login: this server answered the way a classic one does; using the classic login.");
            return null;
        }

        var fields = HotlineHope.AuthenticatedLoginFields(identification, login, password, nickname, iconId, clientVersion);
        if (fields is null)
        {
            DebugLog?.Invoke($"Secure login: the server chose {identification.MacAlgorithm}, which this client can't compute; using the classic login.");
            return null;
        }

        // Advisory only — the login still goes ahead. A NAT in front of a server is ordinary and
        // common; the point is that the operator can see it rather than it passing unnoticed.
        if (identification.EmbeddedEndpoint is { } embedded && _connectedEndpoint is { } actual &&
            (!embedded.Address.Equals(actual.Address) || embedded.Port != actual.Port))
        {
            SecureLoginAddressMismatch?.Invoke(
                $"the server identifies itself as {embedded.Address}:{embedded.Port} but was reached at {actual.Address}:{actual.Port}");
        }

        var loginFields = fields.ToList();
        if (advertiseChatHistory)
        {
            loginFields.Add(new HotlineField(HotlineFieldType.Capabilities, (ushort)HotlineCapabilityBits.ChatHistory));
        }

        var reply = await SendTransactionAsync(HotlineTransactionType.Login, [.. loginFields], ct).ConfigureAwait(false);
        if (reply is not { ErrorCode: 0 })
        {
            LoginRefusedReason = reply?.Field(HotlineFieldType.ErrorText)?.AsString() is { Length: > 0 } text
                ? text
                : "the server refused the login";
            return false;
        }

        SecureLoginAlgorithm = identification.MacAlgorithm;
        ServerApplication = identification.ServerApp;
        DebugLog?.Invoke($"Secure login: authenticated with {identification.MacAlgorithm}" +
            (identification.ServerApp is { Length: > 0 } app ? $" to {app}." : "."));

        await FinishLoginAsync(reply, ct).ConfigureAwait(false);
        return true;
    }

    /// <summary>Classic Hotline's login/password obfuscation — bitwise-NOT every byte (pydora-style "not encryption, just enough to not be plaintext on the wire"). Confirmed against Hotline-Navigator's source: Rust's `!byte` is this exact operation. Public (not just used internally) so it's directly unit-testable without a live server.</summary>
    public static byte[] XorObfuscate(string value)
    {
        var bytes = Encoding.UTF8.GetBytes(value);
        var result = new byte[bytes.Length];
        for (var i = 0; i < bytes.Length; i++)
        {
            result[i] = (byte)~bytes[i];
        }

        return result;
    }

    /// <summary>Best-effort human-readable rendering of one field's raw bytes for DebugLog — as text if it looks printable, a plain number if exactly 2 bytes (most numeric fields are u16), else hex.</summary>
    private static string DescribeField(HotlineField field)
    {
        if (field.Data.Length == 2)
        {
            return field.AsUInt16().ToString();
        }

        var text = field.AsString();
        return text.All(c => !char.IsControl(c)) ? $"\"{text}\"" : Convert.ToHexStringLower(field.Data);
    }

    private uint NextId() => _nextTransactionId++;

    private async Task<HotlineTransactionFrame?> SendTransactionAsync(HotlineTransactionType type, HotlineField[] fields, CancellationToken ct)
    {
        var id = NextId();
        var tcs = new TaskCompletionSource<HotlineTransactionFrame>();
        _pendingReplies[id] = tcs;
        try
        {
            await SendAsync(HotlineTransactionFrame.Create(type, id, fields).Encode(), ct).ConfigureAwait(false);
            return await tcs.Task.WaitAsync(TimeSpan.FromSeconds(10), ct).ConfigureAwait(false);
        }
        catch (TimeoutException)
        {
            return null;
        }
        finally
        {
            _pendingReplies.TryRemove(id, out _);
        }
    }

    protected override int? TryGetFrameLength(IReadOnlyList<byte> buffer)
    {
        if (!_handshakeComplete)
        {
            return buffer.Count >= 8 ? 8 : null;
        }

        return HotlineTransactionFrame.TryGetFrameLength(buffer);
    }

    /// <summary>
    /// One malformed/unexpectedly-shaped incoming transaction (a real risk — see
    /// HotlineUser.Parse's remarks on server variants already found this way) must never take the
    /// whole connection down with it: FramedTcpClient's own receive loop treats ANY exception
    /// escaping PacketReceived as a fatal connection failure and fires Disconnected, tearing the
    /// socket down — which looks exactly like "the server instantly disconnected us" from the
    /// user's side even though the server did nothing wrong. Confirmed necessary live: a real
    /// server (Mobius-based) disconnected a session "after entering chat", not during login/user
    /// list — i.e. from some later, real-time event this code hadn't been exercised against yet.
    /// </summary>
    private void OnPacketReceived(byte[] frame)
    {
        try
        {
            OnPacketReceivedCore(frame);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            ProtocolError?.Invoke(ex);
        }
    }

    private void OnPacketReceivedCore(byte[] frame)
    {
        if (!_handshakeComplete)
        {
            // 8-byte handshake reply: "TRTP" echo (4 bytes) + error code as u32 big-endian.
            var ok = frame.Length >= 8
                && frame.AsSpan(0, 4).SequenceEqual(HotlineConstants.ProtocolId)
                && BinaryPrimitives.ReadUInt32BigEndian(frame.AsSpan(4)) == 0;
            _handshakeComplete = true;
            _handshakeTcs?.TrySetResult(ok);
            return;
        }

        var tx = HotlineTransactionFrame.Decode(frame);
        if (Debug)
        {
            var typeName = Enum.IsDefined(typeof(HotlineTransactionType), tx.Type) ? ((HotlineTransactionType)tx.Type).ToString() : "Unknown";
            var fields = string.Join(", ", tx.Fields.Select(f => $"{(Enum.IsDefined(typeof(HotlineFieldType), f.Type) ? ((HotlineFieldType)f.Type).ToString() : f.Type.ToString())}={DescribeField(f)}"));
            DebugLog?.Invoke($"recv type={typeName}({tx.Type}) id={tx.Id} isReply={tx.IsReply} errorCode={tx.ErrorCode} fields=[{fields}]");
        }

        if (tx.IsReply && _pendingReplies.TryGetValue(tx.Id, out var pending))
        {
            pending.TrySetResult(tx);
            return;
        }

        switch ((HotlineTransactionType)tx.Type)
        {
            case HotlineTransactionType.DisconnectMessage:
                // The real protocol's way of explaining why a server is about to close the
                // connection (kicked, banned, duplicate login, server full, etc.) — confirmed
                // this was silently ignored before, right when a real Mobius-based server was
                // reported disconnecting a session shortly after it entered chat.
                DisconnectMessageReceived?.Invoke(tx.Field(HotlineFieldType.Data)?.AsString() ?? tx.Field(HotlineFieldType.ErrorText)?.AsString() ?? "");
                break;

            case HotlineTransactionType.ChatMessage:
                ChatMessageReceived?.Invoke(tx.Field(HotlineFieldType.Data)?.AsString() ?? "");
                break;

            case HotlineTransactionType.NewMessage:
                if (tx.Field(HotlineFieldType.Data)?.AsString() is { Length: > 0 } posted)
                {
                    FlatNewsPosted?.Invoke(posted);
                }

                break;

            case HotlineTransactionType.UserAccess:
                // Unsolicited, arrives right after connecting — the server telling us our own
                // 64-bit account-access bitmap (guest-level for an anonymous login, richer for a
                // real registered admin/mod account). Confirmed live (observed in Debug logs
                // against multiple real servers) as an 8-byte field 110 payload.
                if (tx.Field(HotlineFieldType.UserAccess) is { Data.Length: 8 } accessField)
                {
                    OwnAccessBits = System.Buffers.Binary.BinaryPrimitives.ReadUInt64BigEndian(accessField.Data);
                }

                break;

            case HotlineTransactionType.ServerMessage:
            {
                // 104 carries two different things. A broadcast from the server has only text; a
                // private message from another user also carries who sent it. Telling them apart by
                // the presence of a sender is what makes a PM a PM — without it every private
                // message showed up as an anonymous "* ..." line in the chat log, with no way to
                // tell who sent it or to reply.
                var text = tx.Field(HotlineFieldType.Data)?.AsString() ?? "";
                var senderName = tx.Field(HotlineFieldType.UserName)?.AsString();
                var senderId = tx.Field(HotlineFieldType.UserId);

                if (senderId is not null || !string.IsNullOrEmpty(senderName))
                {
                    PrivateMessageReceived?.Invoke(new HotlinePrivateMessage(
                        senderId?.AsUInt16() ?? 0,
                        string.IsNullOrEmpty(senderName) ? "(unknown)" : senderName,
                        text));
                }
                else
                {
                    ServerMessageReceived?.Invoke(text);
                }

                break;
            }

            case HotlineTransactionType.ShowAgreement:
            {
                // Just signals arrival here — ConnectAndLoginAsync's own grace-window wait is
                // what actually decides whether/when to send Agreed, so auto-accept and the
                // deferred user-list fetch both happen in one properly-ordered place instead of
                // racing against a fire-and-forget send from this event handler.
                _agreementArrivedTcs?.TrySetResult(true);
                if (!AutoAcceptAgreement)
                {
                    // Never silently agree to a server's rules on the user's behalf — see the
                    // user's own explicit instruction. The UI surfaces this and calls
                    // AcceptAgreementAsync only on an explicit user action.
                    var text = tx.Field(HotlineFieldType.ServerAgreement)?.AsString() ?? "";
                    AgreementReceived?.Invoke(text);
                }

                break;
            }

            case HotlineTransactionType.KeepAlive:
                // No-op — this is a client-to-server ping (see KeepAliveLoopAsync); a server
                // wouldn't normally send one back, but there's nothing harmful in ignoring it if
                // one ever did.
                break;

            case HotlineTransactionType.NotifyUserChange:
            {
                // Individual top-level fields here, NOT a packed UserNameWithInfo blob — confirmed
                // against Hotline-Navigator's real handler code, distinct from GetUserNameList's
                // reply shape (see HotlineUser.Parse).
                var changed = new HotlineUser(
                    tx.Field(HotlineFieldType.UserId)?.AsUInt16() ?? 0,
                    tx.Field(HotlineFieldType.UserIconId)?.AsUInt16() ?? 414,
                    tx.Field(HotlineFieldType.UserFlags)?.AsUInt16() ?? 0,
                    tx.Field(HotlineFieldType.UserName)?.AsString() ?? "");
                _users.RemoveAll(u => u.UserId == changed.UserId);
                _users.Add(changed);
                UserChanged?.Invoke(changed);
                break;
            }

            case HotlineTransactionType.NotifyUserDelete:
            {
                var userId = tx.Field(HotlineFieldType.UserId)?.AsUInt16() ?? 0;
                _users.RemoveAll(u => u.UserId == userId);
                UserLeft?.Invoke(userId);
                break;
            }
        }
    }
}
