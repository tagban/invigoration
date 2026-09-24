using System.Net.WebSockets;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Front;

/// <summary>
/// Drives the Front WebSocket RPC connection: connect, ConnectionService
/// bootstrap, the full AuthenticationServer/AuthenticationClient login
/// sequence (including the external web-auth challenge round trip), and
/// GameUtilities.ProcessClientRequest for the Sunken handoff. Ported from
/// core/src/bgs/client.rs's Client — method IDs, service names, and the
/// authentication dispatch loop's exact acceptance condition are copied
/// directly from that source, not reconstructed from the higher-level
/// protocol docs, since the docs describe the sequence narratively rather
/// than as exact dispatch logic.
///
/// Live-verified end-to-end against a real Battle.net account on
/// 2026-08-22: connect → authenticate (including a real browser-driven web
/// challenge) → GameUtilities.ProcessClientRequest → SunkenClient resume
/// handshake all completed successfully in one continuous session.
/// </summary>
public sealed partial class FrontClient : IAsyncDisposable
{
    /// <summary>Front WebSocket endpoint for the US region — the only region exercised so far. A real client picks this per-account from <see cref="LogonResult.AvailableRegions"/>/<see cref="LogonResult.ConnectedRegion"/>; not implemented yet.</summary>
    public const string DefaultUsUri = "wss://us.actual.battle.net:1119/";

    private const uint ResponseServiceId = 0xfe;
    private const string SubProtocol = "v1.rpc.battle.net";

    private static readonly uint ConnectionServiceHash = ServiceHash.Compute(FrontServices.Connection);
    private static readonly uint AuthenticationServiceHash = ServiceHash.Compute(FrontServices.AuthenticationServer);
    private static readonly uint AuthenticationListenerHash = ServiceHash.Compute(FrontServices.AuthenticationClient);
    private static readonly uint ChallengeListenerHash = ServiceHash.Compute(FrontServices.ChallengeNotify);
    private static readonly uint GameUtilitiesServiceHash = ServiceHash.Compute(FrontServices.GameUtilities);

    private readonly ClientWebSocket _socket = new();

    /// <summary>ClientWebSocket allows one send at a time; Echo replies and requests can now overlap.</summary>
    private readonly SemaphoreSlim _sendLock = new(1, 1);
    private int _lastToken;
    private bool _connected;

    public async Task ConnectAsync(Uri uri, CancellationToken cancellationToken = default)
    {
        _socket.Options.AddSubProtocol(SubProtocol);
        await _socket.ConnectAsync(uri, cancellationToken).ConfigureAwait(false);
    }

    public async Task<ConnectResponse> EstablishAsync(CancellationToken cancellationToken = default)
    {
        await SendAsync(new Header { ServiceId = 0, MethodId = 1, Token = 0, ServiceHash = ConnectionServiceHash },
            new ConnectRequest { UseBindlessRpc = true }.Encode(), cancellationToken).ConfigureAwait(false);
        var (header, body) = await ReceiveAsync(cancellationToken).ConfigureAwait(false);
        RequireResponseOk(header, "ConnectionService.Connect");
        _connected = true;
        return ConnectResponse.Decode(body);
    }

    /// <summary>
    /// Runs the full logon sequence, invoking <paramref name="challengeHandler"/>
    /// (expected to drive a WebView/WebAuthenticationBroker) if and only if
    /// Battle.net issues an external web challenge. Loops until a
    /// LogonResult has arrived and any in-flight web-credential verification
    /// has been acknowledged, matching upstream's exact loop condition.
    /// </summary>
    public async Task<LogonResult> AuthenticateAsync(
        byte[]? cachedWebCredentials,
        Func<Uri, CancellationToken, Task<byte[]>> challengeHandler,
        CancellationToken cancellationToken = default)
    {
        RequireConnected();
        RequireNotListening("AuthenticateAsync");
        var logonToken = await RequestAsync(AuthenticationServiceHash, 1, LogonBuilder.BuildLogonRequest(cachedWebCredentials).Encode(), cancellationToken).ConfigureAwait(false);

        uint? verificationToken = null;
        var logonResponseSeen = false;
        var verificationResponseSeen = false;
        LogonResult? result = null;

        // The web sign-in runs alongside the connection, not instead of it: Battle.net keeps
        // sending Echo keepalives while the user types, and one left unanswered gets the
        // connection dropped, which leaves the sign-in page spinning with nothing to finish.
        Task<(Header Header, byte[] Body)>? receiving = null;
        Task<byte[]>? signingIn = null;

        while (!(logonResponseSeen && result is not null && (verificationToken is null || verificationResponseSeen) && signingIn is null))
        {
            receiving ??= ReceiveApplicationFrameAsync(cancellationToken);
            if (signingIn is not null && await Task.WhenAny(receiving, signingIn).ConfigureAwait(false) == signingIn)
            {
                var credential = await signingIn.ConfigureAwait(false);
                signingIn = null;
                verificationToken = await RequestAsync(AuthenticationServiceHash, 7,
                    new VerifyWebCredentialsRequest { WebCredentials = credential }.Encode(), cancellationToken).ConfigureAwait(false);
                continue;
            }

            var (header, body) = await receiving.ConfigureAwait(false);
            receiving = null;

            if (header.ServiceId == ResponseServiceId)
            {
                if ((header.Status ?? 0) != 0)
                {
                    throw new InvalidOperationException($"Authentication RPC failed with status {header.Status}.");
                }

                if (header.Token == logonToken)
                {
                    logonResponseSeen = true;
                }
                else if (header.Token == verificationToken)
                {
                    verificationResponseSeen = true;
                }

                continue;
            }

            if (header.ServiceHash == AuthenticationListenerHash && header.MethodId == 5)
            {
                result = LogonResult.Decode(body);
            }
            else if (header.ServiceHash == AuthenticationListenerHash && header.MethodId == 10)
            {
                var update = LogonUpdateRequest.Decode(body);
                if (update.ErrorCode != 0)
                {
                    throw new InvalidOperationException($"Battle.net logon update failed: {update.ErrorCode}.");
                }
            }
            else if (header.ServiceHash == AuthenticationListenerHash && header.MethodId == 14)
            {
                var selection = GameAccountSelectedRequest.Decode(body);
                if (selection.Result != 0)
                {
                    throw new InvalidOperationException($"Battle.net game-account selection failed: {selection.Result}.");
                }
            }
            else if (header.ServiceHash == AuthenticationListenerHash && header.MethodId is >= 11 and <= 13)
            {
                // NO_RESPONSE callbacks — nothing to do.
            }
            else if (header.ServiceHash == ChallengeListenerHash && header.MethodId == 3)
            {
                if (verificationToken is not null || signingIn is not null)
                {
                    throw new InvalidOperationException("Battle.net issued more than one web challenge.");
                }

                var challenge = ChallengeExternalRequest.Decode(body);
                signingIn = challengeHandler(challenge.GetValidatedWebAuthUrl(), cancellationToken);
            }
            else if (header.ServiceHash == ChallengeListenerHash && header.MethodId == 4)
            {
                var challengeResult = ChallengeExternalResult.Decode(body);
                if (challengeResult.Passed == false)
                {
                    throw new InvalidOperationException("Battle.net rejected the external challenge.");
                }
            }
            else
            {
                throw new InvalidOperationException($"Unexpected authentication callback service_hash={header.ServiceHash} method={header.MethodId}.");
            }
        }

        return result!;
    }

    public Task<byte[]> GenerateWebCredentialsAsync(CancellationToken cancellationToken = default) =>
        GenerateWebCredentialsAsync("S2", cancellationToken);

    /// <summary>
    /// Asks for a reusable sign-in for <paramref name="program"/> ("S2", "S1", "W3", "OSI", "Fen"):
    /// the "keep me signed in" value a later logon presents as cached_web_credentials. Battle.net
    /// issues one per game, so a single signed-in session can collect one for each.
    /// </summary>
    public async Task<byte[]> GenerateWebCredentialsAsync(string program, CancellationToken cancellationToken = default)
    {
        var body = await CallAsync(AuthenticationServiceHash, 8, new GenerateWebCredentialsRequest { Program = FourCc.Encode(program) }.Encode(),
            "AuthenticationService.GenerateWebCredentials", cancellationToken).ConfigureAwait(false);
        var response = GenerateWebCredentialsResponse.Decode(body);
        return response.WebCredentials ?? throw new InvalidOperationException("GenerateWebCredentials returned no credential.");
    }

    public async Task<SunkenHandoff> ProcessClientRequestAsync(EntityId gameAccountId, byte[] sessionKey, CancellationToken cancellationToken = default)
    {
        var request = GameUtilitiesSession.BuildProcessClientRequest(gameAccountId, sessionKey);
        var body = await CallAsync(GameUtilitiesServiceHash, 1, request.Encode(), "GameUtilities.ProcessClientRequest", cancellationToken).ConfigureAwait(false);
        return GameUtilitiesSession.ParseHandoff(ClientResponse.Decode(body));
    }

    public async Task CloseAsync(CancellationToken cancellationToken = default)
    {
        _connected = false;
        _closing = true;
        if (_socket.State == WebSocketState.Open)
        {
            await _socket.CloseAsync(WebSocketCloseStatus.NormalClosure, null, cancellationToken).ConfigureAwait(false);
        }
    }

    private async Task<uint> RequestAsync(uint serviceHash, uint method, byte[] body, CancellationToken cancellationToken)
    {
        var token = NextToken();
        await SendAsync(new Header { ServiceId = 0, MethodId = method, Token = token, ServiceHash = serviceHash }, body, cancellationToken).ConfigureAwait(false);
        return token;
    }

    private Task RespondAsync(uint token, byte[] body, CancellationToken cancellationToken) =>
        SendAsync(new Header { ServiceId = ResponseServiceId, Token = token, Status = 0 }, body, cancellationToken);

    private async Task SendAsync(Header header, byte[] body, CancellationToken cancellationToken)
    {
        var frame = FrontFrame.Encode(header, body);
        await _sendLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            await _socket.SendAsync(frame, WebSocketMessageType.Binary, endOfMessage: true, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _sendLock.Release();
        }
    }

    /// <summary>Reads exactly one Front RPC message: one binary WebSocket message, reassembled across however many fragments the transport chose to split it into.</summary>
    private async Task<(Header Header, byte[] Body)> ReceiveAsync(CancellationToken cancellationToken)
    {
        var buffer = new List<byte>();
        var chunk = new byte[8192];
        while (true)
        {
            var result = await _socket.ReceiveAsync(chunk, cancellationToken).ConfigureAwait(false);
            if (result.MessageType == WebSocketMessageType.Close)
            {
                throw new InvalidOperationException("Front WebSocket closed by the server.");
            }

            buffer.AddRange(chunk.AsSpan(0, result.Count).ToArray());
            if (result.EndOfMessage)
            {
                break;
            }
        }

        return FrontFrame.Decode(buffer.ToArray());
    }

    /// <summary>Like <see cref="ReceiveAsync"/>, but transparently answers ConnectionService.Echo (a keepalive callback) instead of surfacing it, matching upstream's automatic-Echo behavior.</summary>
    private async Task<(Header Header, byte[] Body)> ReceiveApplicationFrameAsync(CancellationToken cancellationToken)
    {
        while (true)
        {
            var (header, body) = await ReceiveAsync(cancellationToken).ConfigureAwait(false);
            if (header.ServiceHash == ConnectionServiceHash && header.MethodId == 3 && header.ServiceId != ResponseServiceId)
            {
                var echo = EchoRequest.Decode(body);
                await RespondAsync(header.Token, new EchoResponse { Time = echo.Time, Payload = echo.Payload }.Encode(), cancellationToken).ConfigureAwait(false);
                continue;
            }

            return (header, body);
        }
    }

    private async Task<(Header Header, byte[] Body)> AwaitResponseAsync(uint token, string operation, CancellationToken cancellationToken)
    {
        var (header, body) = await ReceiveApplicationFrameAsync(cancellationToken).ConfigureAwait(false);
        if (header.ServiceId != ResponseServiceId)
        {
            throw new InvalidOperationException($"Unexpected callback during {operation}: service_hash={header.ServiceHash} method={header.MethodId}.");
        }

        if (header.Token != token)
        {
            throw new InvalidOperationException($"Unexpected response token during {operation}: {header.Token}.");
        }

        RequireResponseOk(header, operation);
        return (header, body);
    }

    private static void RequireResponseOk(Header header, string operation)
    {
        if (header.ServiceId != ResponseServiceId)
        {
            throw new InvalidOperationException($"Expected a response for {operation} but got a callback (service_hash={header.ServiceHash}, method={header.MethodId}).");
        }

        if ((header.Status ?? 0) != 0)
        {
            throw new InvalidOperationException($"{operation} failed with status {header.Status}.");
        }
    }

    private void RequireConnected()
    {
        if (!_connected)
        {
            throw new InvalidOperationException("Front connection has not been established yet.");
        }
    }

    public async ValueTask DisposeAsync()
    {
        _closing = true;
        if (_socket.State is WebSocketState.Open or WebSocketState.CloseReceived)
        {
            try
            {
                await _socket.CloseAsync(WebSocketCloseStatus.NormalClosure, null, CancellationToken.None).ConfigureAwait(false);
            }
            catch (WebSocketException)
            {
                // Best-effort close during disposal.
            }
        }

        _socket.Dispose();
        if (_listenTask is not null)
        {
            try
            {
                await _listenTask.ConfigureAwait(false);
            }
            catch
            {
                // The listener reports its own end through ListeningStopped.
            }
        }
    }
}
