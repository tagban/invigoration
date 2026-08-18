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

    private readonly BncsConnection _bncs = new();
    private readonly BnlsConnection _bnls = new();
    private readonly RealmConnection _realm = new();
    private readonly AuthState _auth = new();
    private readonly BotSessionState _session = new();
    private DateTimeOffset _connectedAt;

    public BotConfig Config { get; }

    /// <summary>When on, logs every raw BNCS/BNLS packet sent and received as a hex dump.</summary>
    public bool DebugMode
    {
        get => _session.DebugMode;
        set => _session.DebugMode = value;
    }

    public event Action<IReadOnlyList<ChatLogSegment>>? Log;
    public event Action? BnlsConnected;
    public event Action? BncsConnected;
    public event Action<Exception?>? BncsDisconnected;
    public event Action<IReadOnlyList<string>>? ChannelListReceived;

    public event Action<ChatEvent>? ChatMessage;

    public BotEngine(BotConfig config)
    {
        Config = config;

        _bnls.Connected += OnBnlsConnected;
        _bnls.PacketReceived += frame => SafeFireAndForget(HandleBnlsPacket(frame), "handling a BNLS packet");
        _bnls.Disconnected += ex => LogInfo($"BNLS connection closed{(ex is null ? "." : $": {ex.Message}")}");

        _bncs.Connected += OnBncsConnected;
        _bncs.PacketReceived += frame => SafeFireAndForget(HandleBncsPacket(frame), "handling a BNCS packet");
        _bncs.Disconnected += ex =>
        {
            LogError($"Battle.net disconnected{(ex is null ? "." : $": {ex.Message}")}");
            BncsDisconnected?.Invoke(ex);
        };
    }

    public async Task ConnectAsync(CancellationToken cancellationToken = default)
    {
        if (Config.ConnectionMode == ConnectionMode.TelnetGateway)
        {
            throw new NotSupportedException(
                "Telnet/chat-gateway mode is planned but not implemented yet; use BncsBinary.");
        }

        if (BncsProduct.IsLikelyIncompatible(Config.Product, Config.BattlenetServer))
        {
            LogWarning(
                $"{BncsProduct.GetDisplayName(Config.Product)} was retired from official Battle.net and will " +
                "likely be rejected by this server; it still works against PVPGN/private servers.");
        }

        LogInfo($"Battle.net Login Server connecting to {Config.BnlsServer}...");
        await _bnls.ConnectAsync(Config.BnlsServer, Config.BnlsPort, cancellationToken).ConfigureAwait(false);
    }

    public Task DisconnectAsync()
    {
        _bncs.Close();
        _bnls.Close();
        _realm.Close();
        LogInfo("Disconnected.");
        return Task.CompletedTask;
    }

    public Task SendChatCommandAsync(string text) =>
        SendBncsAsync(new PacketWriter().WriteNTString(text), BncsPacketId.SID_CHATCOMMAND);

    public async Task JoinHomeAsync()
    {
        await SendBncsAsync(new PacketWriter(), BncsPacketId.SID_LEAVECHAT).ConfigureAwait(false);
        await SendBncsAsync(
            new PacketWriter().WriteDword(2).WriteNTString(Config.HomeChannel),
            BncsPacketId.SID_JOINCHANNEL).ConfigureAwait(false);
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

    private void LogInfo(string message) => LogLine(new ChatLogSegment(ChatColors.Green, message));

    private void LogWarning(string message) => LogLine(new ChatLogSegment(ChatColors.Orange, message));

    private void LogError(string message) => LogLine(new ChatLogSegment(ChatColors.Red, message));

    private void LogDebug(string message)
    {
        if (_session.DebugMode)
        {
            LogLine(new ChatLogSegment(ChatColors.HexPink, message));
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

    public ValueTask DisposeAsync()
    {
        _bncs.Close();
        _bnls.Close();
        _realm.Close();
        return ValueTask.CompletedTask;
    }
}
