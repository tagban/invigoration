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

    public bool HasOwnAccess(int bit) => (OwnAccessBits & (1UL << bit)) != 0;

    /// <summary>Off by default — never silently agree to a server's rules on the user's behalf. Set before ConnectAndLoginAsync; per-tracker, from HotlineTrackerConfig.AutoAcceptAgreement.</summary>
    public bool AutoAcceptAgreement { get; set; }

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
            return false;
        }

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
                ServerMessageReceived?.Invoke(tx.Field(HotlineFieldType.Data)?.AsString() ?? "");
                break;

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
