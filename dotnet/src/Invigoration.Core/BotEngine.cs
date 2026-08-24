using Invigoration.Core.Auth;
using Invigoration.Core.Chat;
using Invigoration.Core.Commands;
using Invigoration.Core.Config;
using Invigoration.Core.Networking;
using Invigoration.Core.Protocol;

namespace Invigoration.Core;

/// <summary>
/// Orchestrates one bot's BNCS + BNLS (+ D2 realm) connections and the login
/// handshake between them. Replaces frmMain's Winsock event handlers and the
/// global-state ParseBnet/ParseBNLS dispatch in modBNET.bas/modBNLS.bas. One
/// instance per bot tab.
/// </summary>
public sealed partial class BotEngine : IAsyncDisposable
{
    private const string BnlsClientName = "Invigoration";
    private const int RealmPort = 6112;

    /// <summary>
    /// How often to proactively send SID_NULL while connected, matching the
    /// VB6 original's tmrAntiIdle (800ms ticks, firing at count 110). Official
    /// Battle.net can silently idle-drop a connection that goes quiet even
    /// while still answering SID_PING, so a real client sends this too.
    /// </summary>
    private static readonly TimeSpan KeepAliveInterval = TimeSpan.FromSeconds(88);

    private readonly BncsConnection _bncs = new();
    private readonly BnlsConnection _bnls = new();
    private readonly RealmConnection _realm = new();
    private readonly AuthState _auth = new();
    private readonly BotSessionState _session = new();
    private readonly List<FriendEntry> _friends = [];
    private DateTimeOffset _connectedAt;
    private CancellationTokenSource? _keepAliveCts;

    /// <summary>
    /// Process-wide (not per-instance): flood protection needs to hold even
    /// when several of the user's own bots — on different servers or the
    /// same one — send around the same moment, since a per-IP flood
    /// detector (common on PVPGN) can still see that as a burst even though
    /// each individual connection is politely spaced on its own. Learned the
    /// hard way — a per-engine gate alone wasn't enough and still got a test
    /// account flood-banned on a second server during a linked-trivia round.
    /// </summary>
    private static readonly SemaphoreSlim ChatSendGate = new(1, 1);
    private static DateTime _nextChatSendAllowedUtc = DateTime.MinValue;

    /// <summary>
    /// Settable so a UI can swap in an edited config after the user saves
    /// changes to an already-added bot (the engine reads Config.* fresh on
    /// every use rather than caching values, so this is safe at any time).
    /// </summary>
    public BotConfig Config { get; set; }

    /// <summary>When on, logs every raw BNCS/BNLS packet sent and received as a hex dump.</summary>
    public bool DebugMode
    {
        get => _session.DebugMode;
        set => _session.DebugMode = value;
    }

    /// <summary>The active named color set (StarCraft/Diablo II/Warcraft III) this bot's chat log renders with.</summary>
    public ChatPalette Palette => ChatPalette.ForScheme(Config);

    public event Action<IReadOnlyList<ChatLogSegment>>? Log;

    /// <summary>
    /// A non-command message this bot itself just sent, echoed locally since Battle.net never
    /// echoes a client's own outgoing channel messages back as a real chat event — see
    /// SendChatCommandAsync. Deliberately its own event rather than folded into the generic Log
    /// above: the one existing Log subscriber (BotTabViewModel.OnLog) always writes into the flat
    /// ChatLines collection, which is hidden entirely for a SupportsMultiChannel (SC2/SC:R/WC3:R)
    /// bot — a self-sent message routed that way was invisible on those bots, not just missing
    /// the same speaker icon a real Talk event gets. This lets the App layer route it to wherever
    /// the message actually went (the active sub-tab) and resolve an icon for it the same way.
    /// </summary>
    public event Action<IReadOnlyList<ChatLogSegment>>? SelfChatSent;
    public event Action? BnlsConnected;
    public event Action? BncsConnected;
    public event Action<Exception?>? BncsDisconnected;
    public event Action<IReadOnlyList<string>>? ChannelListReceived;
    public event Action<IReadOnlyList<FriendEntry>>? FriendsListUpdated;

    public event Action<ChatEvent>? ChatMessage;

    /// <summary>
    /// Raised when BotEngine has mutated its own Config in a way that must
    /// reach bots.json even though nothing routed through the Config window
    /// (currently just BattlenetCredentialProfileId auto-assignment on first
    /// SC2 connect — see BotEngine.Sc2.cs's EnsureBattlenetCredentialProfileId).
    /// Nothing today calls SaveAll() on a bare Connect (only add/remove-bot,
    /// a config-window save, or app close do), so without this an
    /// auto-connect-on-startup bot that never touches those could lose the
    /// assignment and recreate an orphaned profile next launch. The App
    /// layer should treat this exactly like "please save the bot list now".
    /// </summary>
    public event Action? ConfigPersistNeeded;

    public BotEngine(BotConfig config)
    {
        Config = config;
        Trivia.TriviaGroupRegistry.RegisterEngine(this);

        _bnls.Connected += OnBnlsConnected;
        _bnls.PacketReceived += frame => SafeFireAndForget(HandleBnlsPacket(frame), "handling a BNLS packet");
        _bnls.Disconnected += OnBnlsDisconnected;

        _bncs.Connected += OnBncsConnected;
        _bncs.PacketReceived += frame => SafeFireAndForget(HandleBncsPacket(frame), "handling a BNCS packet");
        _bncs.Disconnected += ex =>
        {
            StopKeepAlive();
            _friends.Clear();
            LogError($"Battle.net disconnected{(ex is null ? "." : $": {ex.Message}")}");
            BncsDisconnected?.Invoke(ex);
            MaybeScheduleAutoReconnect();
        };

        WireDiscordBridge();
    }

    private bool _isIntentionalDisconnect;
    private CancellationTokenSource? _autoReconnectCts;

    /// <summary>
    /// Reconnects after an unexpected drop (never after DisconnectAsync was
    /// called deliberately), so the bot recovers on its own from anything
    /// that closes the connection out from under it — a network hiccup, a
    /// server-side idle/session policy, or (on official Battle.net) the
    /// server enforcing SID_REQUIREDWORK/ExtraWork compliance, which this
    /// client doesn't and won't implement (see BotEngine.Bncs.cs remarks on
    /// HandleRequiredWork). This just re-establishes a connection the way
    /// any client would after losing its socket — it doesn't change what
    /// happens once connected, so it doesn't paper over that specifically.
    /// Off by default: repeated automatic reconnects against a server that's
    /// actively dropping this client for policy reasons could look more
    /// automated, not less, so this is opt-in rather than a default-on
    /// resilience feature.
    /// </summary>
    private void MaybeScheduleAutoReconnect()
    {
        _autoReconnectCts?.Cancel();

        if (_isIntentionalDisconnect || !Config.AutoReconnect)
        {
            return;
        }

        var delaySeconds = Math.Max(1, Config.AutoReconnectDelaySeconds);
        _autoReconnectCts = new CancellationTokenSource();
        var token = _autoReconnectCts.Token;
        SafeFireAndForget(RunAutoReconnectAsync(delaySeconds, token), "auto-reconnecting");
    }

    private async Task RunAutoReconnectAsync(int delaySeconds, CancellationToken cancellationToken)
    {
        LogInfo($"Unexpected disconnect — reconnecting in {delaySeconds}s (auto-reconnect is on).");
        await Task.Delay(TimeSpan.FromSeconds(delaySeconds), cancellationToken).ConfigureAwait(false);
        await ConnectAsync(cancellationToken).ConfigureAwait(false);
    }

    private void StartKeepAlive()
    {
        StopKeepAlive();
        _keepAliveCts = new CancellationTokenSource();
        SafeFireAndForget(RunKeepAliveLoopAsync(_keepAliveCts.Token), "sending a keep-alive");
    }

    private void StopKeepAlive()
    {
        _keepAliveCts?.Cancel();
        _keepAliveCts?.Dispose();
        _keepAliveCts = null;
    }

    private async Task RunKeepAliveLoopAsync(CancellationToken cancellationToken)
    {
        try
        {
            using var timer = new PeriodicTimer(KeepAliveInterval);
            while (await timer.WaitForNextTickAsync(cancellationToken).ConfigureAwait(false))
            {
                await SendBncsAsync(new PacketWriter(), BncsPacketId.SID_NULL).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException)
        {
            // Normal on disconnect — StopKeepAlive() cancels this loop deliberately.
        }
    }

    public async Task ConnectAsync(CancellationToken cancellationToken = default)
    {
        _isIntentionalDisconnect = false;
        StartDiscordBridgeIfEnabled();

        if (BncsProduct.IsStimpakBacked(Config.Product))
        {
            await ConnectSc2Async(cancellationToken).ConfigureAwait(false);
            return;
        }

        if (BncsProduct.IsLikelyIncompatible(Config.Product, Config.BattlenetServer))
        {
            LogWarning(
                $"{BncsProduct.GetDisplayName(Config.Product)} was retired from official Battle.net and will " +
                "likely be rejected by this server; it still works against PVPGN/private servers.");
        }

        LogInfo($"Battle.net Login Server connecting to {Config.BnlsServer}...");
        await _bnls.ConnectAsync(Config.BnlsServer, Config.BnlsPort, cancellationToken, BuildProxyOptions()).ConfigureAwait(false);
    }

    /// <summary>Null unless the user turned proxying on for this bot — passed through to every ConnectAsync call so BNCS/BNLS/realm all tunnel through the same proxy.</summary>
    private ProxyOptions? BuildProxyOptions() => Config.ProxyEnabled
        ? new ProxyOptions(
            Config.ProxyProtocol,
            Config.ProxyHost,
            Config.ProxyPort,
            string.IsNullOrEmpty(Config.ProxyUsername) ? null : Config.ProxyUsername,
            string.IsNullOrEmpty(Config.ProxyUsername) ? null : Config.ProxyPassword)
        : null;

    public async Task DisconnectAsync()
    {
        _isIntentionalDisconnect = true;
        _autoReconnectCts?.Cancel();
        _bncs.Close();
        _bnls.Close();
        _realm.Close();
        await DisconnectSc2Async().ConfigureAwait(false);
        await StopDiscordBridgeAsync().ConfigureAwait(false);
        LogInfo("Disconnected.");
    }

    /// <summary>
    /// Sends chat text (or a raw "/command" passthrough). Battle.net never
    /// echoes a client's own outgoing channel messages back as a chat event,
    /// so — matching the original modFunctions.bas Send() — this echoes
    /// non-command text into the local log itself; without it, every command
    /// reply (uptime, fudd/canada confirmations, etc.) would be invisible to
    /// the bot's own operator despite other users seeing it fine.
    ///
    /// Every send is flood-protected: a burst (e.g. trivia asking a question,
    /// then almost immediately announcing a fast correct answer) is spaced
    /// out to at least Config.FloodProtectionDelayMs apart rather than fired
    /// back-to-back, since neither this port nor the original VB6 bot ever
    /// had any throttling and Battle.net/PVPGN servers will disconnect or
    /// mute a client that sends too many lines too quickly. The gate is
    /// shared across every bot in the process (see ChatSendGate), not just
    /// this connection, so several of the user's own linked bots sending
    /// around the same moment still queue up rather than bursting together.
    /// </summary>
    public async Task SendChatCommandAsync(string text, byte? sc2ChannelOverride = null)
    {
        await ChatSendGate.WaitAsync().ConfigureAwait(false);
        try
        {
            var waitMs = (_nextChatSendAllowedUtc - DateTime.UtcNow).TotalMilliseconds;
            if (waitMs > 0)
            {
                await Task.Delay((int)waitMs).ConfigureAwait(false);
            }

            _nextChatSendAllowedUtc = DateTime.UtcNow.AddMilliseconds(Math.Max(0, Config.FloodProtectionDelayMs));

            var isSlashCommand = text.Length > 0 && text[0] == '/';
            var outgoing = isSlashCommand ? text : ApplyTextEffects(text);

            if (BncsProduct.IsStimpakBacked(Config.Product))
            {
                await SendSc2Async(outgoing, sc2ChannelOverride).ConfigureAwait(false);
            }
            else
            {
                await SendBncsAsync(new PacketWriter().WriteNTString(outgoing), BncsPacketId.SID_CHATCOMMAND)
                    .ConfigureAwait(false);
            }

            // Stimpak-backed (SC2/SC:R/WC3:R) products don't need this local echo at all — unlike
            // classic BNCS, Stimpak's own protocol genuinely echoes a sent channel message back
            // through the normal event stream (MessageReceived), the same way it already does for
            // a sent whisper (WhisperReceived{Outgoing:true} — see BotEngine.Sc2.cs). Echoing it
            // here too doubled every SC2 message: once here (with no real BattleTag to show, since
            // Config.Username is a classic-BNCS-only field — Stimpak logs in via OAuth, not a
            // username/password Config ever populates), and once for real once the server's own
            // echo arrived with the correct name and clan tag.
            if (!isSlashCommand && !BncsProduct.IsStimpakBacked(Config.Product))
            {
                var segments = new List<ChatLogSegment> { new(Palette.SelfUserName, $"{Config.Username}: ") };
                segments.AddRange(ChatColorFormatter.Parse(outgoing, Palette.White, Palette));
                SelfChatSent?.Invoke(segments);
            }
        }
        finally
        {
            ChatSendGate.Release();
        }
    }

    /// <summary>
    /// The joke text-transform modes toggled by the "fudd" and "canada"
    /// commands — applied to every outgoing chat message (not raw "/"
    /// commands) while active, which is what makes turning them on visibly
    /// change the bot's own replies too.
    /// </summary>
    private string ApplyTextEffects(string text)
    {
        if (_session.FuddMode)
        {
            text = text.Replace('r', 'w').Replace('R', 'W');
        }

        if (_session.CanadaMode)
        {
            text += ", eh?";
        }

        return text;
    }

    public async Task JoinHomeAsync()
    {
        await SendBncsAsync(new PacketWriter(), BncsPacketId.SID_LEAVECHAT).ConfigureAwait(false);
        await SendBncsAsync(
            new PacketWriter().WriteDword(2).WriteNTString(Config.HomeChannel),
            BncsPacketId.SID_JOINCHANNEL).ConfigureAwait(false);
    }

    /// <summary>
    /// Re-requests the full friends list from the server. Sent automatically
    /// once on entering chat; also exposed for manual refresh, since Diablo
    /// II's server doesn't push SID_FRIENDSUPDATE automatically and needs
    /// polling to pick up status changes (per bnetdocs.org).
    /// </summary>
    public Task RequestFriendsListAsync() => SendBncsAsync(new PacketWriter(), BncsPacketId.SID_FRIENDSLIST);

    /// <summary>
    /// BNLS is only needed transiently, for the version/CD-key/password
    /// hashing steps of the login handshake — the server itself closes the
    /// connection once that's done, which is expected and not worth
    /// surfacing to the operator as if something went wrong. Only log it
    /// when it happens before BNCS logon finished (a real failure) or when
    /// the close carried an actual exception.
    /// </summary>
    private void OnBnlsDisconnected(Exception? ex)
    {
        if (ex is null && _auth.LoggedOnToBncs)
        {
            LogDebug("BNLS connection closed.");
            return;
        }

        LogInfo($"BNLS connection closed{(ex is null ? "." : $": {ex.Message}")}");
    }

    private async void OnBnlsConnected()
    {
        try
        {
            LogInfo("Battle.net Login Server connected!");
            BnlsConnected?.Invoke();
            await SendBnlsAsync(new PacketWriter().WriteNTString(BnlsClientName), BnlsPacketId.BNLS_AUTHORIZE)
                .ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            LogError($"Error while starting the BNLS handshake: {ex.Message}");
        }
    }

    private async void OnBncsConnected()
    {
        try
        {
            LogInfo("Battle.net Connected!");
            BncsConnected?.Invoke();
            StartKeepAlive();
            await _bncs.SendAsync([0x01]).ConfigureAwait(false); // BNCS binary-protocol byte
            await SendAuthInfoAsync().ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            LogError($"Error while starting the BNCS handshake: {ex.Message}");
        }
    }

    private async Task SendAuthInfoAsync()
    {
        var writer = new PacketWriter()
            .WriteDword(0) // Protocol ID
            .WriteAscii("68XI") // Platform ID "IX86", stored wire-reversed
            .WriteAscii(Config.Product) // Product ID, already stored wire-reversed
            .WriteDword(_auth.VersionByte)
            .WriteDword(0) // Product language
            .WriteDword(0) // Local IP
            .WriteDword(0x480) // Time zone bias
            .WriteDword(0x409) // Locale ID (en-US)
            .WriteDword(0x1033) // Language ID (en-US)
            .WriteNTString("USA")
            .WriteNTString("United States");
        await SendBncsAsync(writer, BncsPacketId.SID_AUTH_INFO).ConfigureAwait(false);

        if (Config.ZeroPing)
        {
            // Fabricate one fast ping response now; the SID_PING handler then
            // stops responding entirely so the server can't recalculate it.
            await SendBncsAsync(new PacketWriter().WriteDword(0), BncsPacketId.SID_PING).ConfigureAwait(false);
        }
    }

    private Task SendPasswordHashRequestAsync(string password)
    {
        var writer = new PacketWriter()
            .WriteDword((uint)password.Length)
            .WriteDword(0)
            .WriteAscii(password);
        return SendBnlsAsync(writer, BnlsPacketId.BNLS_HASHDATA);
    }

    private Task SendBncsAsync(PacketWriter writer, BncsPacketId id)
    {
        var packet = writer.ToBncsPacket(id);
        LogDebug($"BNCS send 0x{(byte)id:X2} ({id}), {packet.Length} bytes: {ToHexDump(packet)}");
        return _bncs.SendAsync(packet);
    }

    private Task SendBnlsAsync(PacketWriter writer, BnlsPacketId id)
    {
        var packet = writer.ToBnlsPacket(id);
        LogDebug($"BNLS send 0x{(byte)id:X2} ({id}), {packet.Length} bytes: {ToHexDump(packet)}");
        return _bnls.SendAsync(packet);
    }

    private void LogLine(params ChatLogSegment[] segments) => Log?.Invoke(segments);

    private void LogInfo(string message) => LogLine(new ChatLogSegment(Palette.Info, message));

    private void LogWarning(string message) => LogLine(new ChatLogSegment(Palette.Debug, message));

    private void LogError(string message) => LogLine(new ChatLogSegment(Palette.Error, message));

    private void LogDebug(string message)
    {
        if (_session.DebugMode)
        {
            // Timestamped (unlike the other Log* helpers) specifically so packet
            // timing questions — "is the server pinging too often, or does it just
            // look that way from a wall of untimed lines?" — can be answered by
            // reading the log instead of guessing from line order.
            LogLine(new ChatLogSegment(Palette.Debug, $"[{DateTime.Now:HH:mm:ss.fff}] {message}"));
        }
    }

    private static string ToHexDump(byte[] data) => Convert.ToHexString(data);

    /// <summary>
    /// Awaits a fire-and-forget async handler and logs any exception instead
    /// of letting it vanish as an unobserved task exception — without this,
    /// a throwing packet handler just silently stops the handshake dead with
    /// no visible error, which is exactly what happened before this existed.
    /// </summary>
    private async void SafeFireAndForget(Task task, string context)
    {
        try
        {
            await task.ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            LogError($"Error while {context}: {ex.Message}");
            LogDebug(ex.ToString());
        }
    }

    public async ValueTask DisposeAsync()
    {
        Trivia.TriviaGroupRegistry.UnregisterEngine(this);
        _bncs.Close();
        _bnls.Close();
        _realm.Close();
        await DisconnectSc2Async().ConfigureAwait(false);
        await StopDiscordBridgeAsync().ConfigureAwait(false);
    }
}
