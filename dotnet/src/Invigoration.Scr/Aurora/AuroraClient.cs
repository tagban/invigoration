using System.Net.WebSockets;
using System.Text;
using System.Text.Json.Nodes;
using System.Threading.Channels;
using Invigoration.Sc2.Protobuf;

namespace Invigoration.Scr.Aurora;

/// <summary>What Aurora's logon hands over: the session key and IDs the classic sign-in needs.</summary>
public sealed record AuroraSession(byte[] SessionKey, ulong AccountHigh, ulong AccountLow, ulong GameAccountHigh, ulong GameAccountLow, ulong ConnectedRegion, string? BattleTag);

/// <summary>Where the classic chat server is, and the one-time ticket to present there.</summary>
public sealed record ClassicEndpoint(Uri Url, byte[] Ticket);

/// <summary>
/// SC:R's Aurora connection: the same Battle.net RPC as SC2's Front (connect, logon, web
/// challenge, VerifyWebCredentials, ProcessClientRequest), carried as JSON [header, body] pairs
/// over wss://us.actual.battle.net:1119/ with subprotocol jsonrpc.aurora.v1.30.battle.net.
/// Service hashes and bodies as in ncarrillo/sc1-research (MIT), and our own captures.
/// It keeps answering Connection's Echo the whole time, the web sign-in included: unanswered,
/// Battle.net drops the connection after a couple of minutes.
/// </summary>
public sealed class AuroraClient : IAsyncDisposable
{
    public static readonly Uri UsEndpoint = new("wss://us.actual.battle.net:1119/");
    public const string SubProtocol = "jsonrpc.aurora.v1.30.battle.net";

    public const uint ConnectionService = 0x65446991;
    public const uint AuthenticationService = 0x0DECFC01;
    public const uint GameUtilitiesService = 0x3FC1274D;
    public const uint AuthenticationListener = 0x71240E35;
    public const uint ChallengeListener = 0xBBDA171F;

    /// <summary>Non-secret stand-in for a saved sign-in: it asks Battle.net for a fresh web challenge.</summary>
    public const string NoSavedCredential = "US-00000000000000000000000000000000-0000";

    private readonly ClientWebSocket _socket = new();
    private readonly SemaphoreSlim _sendLock = new(1, 1);
    private readonly Dictionary<ulong, TaskCompletionSource<JsonArray>> _pending = new();
    private readonly Channel<JsonArray> _calls = Channel.CreateUnbounded<JsonArray>();
    private readonly CancellationTokenSource _stop = new();
    private readonly Action<string> _trace;
    private ulong _nextToken = 1;
    private Task? _reading;

    public AuroraClient(Action<string>? trace = null) => _trace = trace ?? (_ => { });

    public async Task ConnectAsync(Uri endpoint, CancellationToken cancellationToken)
    {
        _socket.Options.AddSubProtocol(SubProtocol);
        await _socket.ConnectAsync(endpoint, cancellationToken).ConfigureAwait(false);
        _reading = Task.Run(ReadLoopAsync);
        await CallAsync(ConnectionService, 1, new JsonObject { ["use_bindless_rpc"] = true }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Logs on with a saved sign-in, or, when there's none or Battle.net turns it down, through
    /// <paramref name="challenge"/>, which shows Battle.net's sign-in page and returns its ST value.
    /// </summary>
    public async Task<AuroraSession> LogonAsync(string program, byte[]? savedCredential, Func<Uri, CancellationToken, Task<byte[]>> challenge, CancellationToken cancellationToken)
    {
        var logon = SendAsync(AuthenticationService, 1, new JsonObject
        {
            ["allow_logon_queue_notifications"] = true,
            ["application_version"] = ScrProtocol.ApplicationVersion,
            ["cached_web_credentials"] = Convert.ToBase64String(savedCredential ?? Encoding.ASCII.GetBytes(NoSavedCredential)),
            ["locale"] = "enUS",
            ["platform"] = "Mac",
            ["program"] = program,
            ["version"] = ScrProtocol.ApplicationVersion.ToString(),
            ["web_client_verification"] = true,
        }, cancellationToken);

        AuroraSession? session = null;
        Task<JsonArray>? verification = null;
        Task<byte[]>? signingIn = null;
        Task<JsonArray>? nextCall = null;
        while (session is null || !logon.IsCompleted || verification is { IsCompleted: false } || signingIn is not null)
        {
            // One read outstanding at a time, kept across passes, so no call is ever lost.
            nextCall ??= _calls.Reader.ReadAsync(cancellationToken).AsTask();
            var waits = new List<Task> { nextCall };
            if (!logon.IsCompleted)
            {
                waits.Add(logon);
            }

            if (signingIn is not null)
            {
                waits.Add(signingIn);
            }

            if (verification is { IsCompleted: false })
            {
                waits.Add(verification);
            }

            var done = await Task.WhenAny(waits).ConfigureAwait(false);
            if (done == signingIn)
            {
                var credential = await signingIn.ConfigureAwait(false);
                signingIn = null;
                verification = SendAsync(AuthenticationService, 7, new JsonObject { ["web_credentials"] = Convert.ToBase64String(credential) }, cancellationToken);
                continue;
            }

            if (done != nextCall)
            {
                await (Task)done; // surfaces a failed logon or verification
                continue;
            }

            var call = await nextCall.ConfigureAwait(false);
            nextCall = null;
            var header = call[0]!.AsObject();
            var body = call[1]!.AsObject();
            var service = Number(header["service_hash"]);
            var method = Number(header["method_id"]);
            if (service == AuthenticationListener && method == 5)
            {
                session = ParseLogonResult(body);
                _trace($"Aurora logon complete: {session.BattleTag ?? "account"} in region {session.ConnectedRegion}.");
            }
            else if (service == ChallengeListener && method == 3)
            {
                if (signingIn is not null || verification is not null)
                {
                    throw new InvalidOperationException("Battle.net asked for more than one web sign-in.");
                }

                var url = ParseChallengeUrl(body);
                _trace($"Battle.net asked for a web sign-in at {url.Host}{url.AbsolutePath}.");
                signingIn = challenge(url, cancellationToken);
            }
            else if (service == ChallengeListener && method == 4 && body["passed"] is JsonValue passed && !IsTrue(passed))
            {
                throw new InvalidOperationException("Battle.net rejected the web sign-in.");
            }
        }

        await logon.ConfigureAwait(false);
        if (verification is not null)
        {
            await verification.ConfigureAwait(false);
        }

        return session;
    }

    /// <summary>A reusable "keep me signed in" credential for <paramref name="program"/>, issued now that we're signed in.</summary>
    public async Task<byte[]?> GenerateWebCredentialsAsync(string program, CancellationToken cancellationToken)
    {
        var reply = await CallAsync(AuthenticationService, 8, new JsonObject { ["program"] = FourCc(program) }, cancellationToken).ConfigureAwait(false);
        return reply[1]?["web_credentials"]?.GetValue<string>() is { Length: > 0 } text ? Convert.FromBase64String(text) : null;
    }

    /// <summary>GameUtilities.ProcessClientRequest with a ConnectToServerRequest: where the classic chat server is.</summary>
    public async Task<ClassicEndpoint> ConnectToServerAsync(CancellationToken cancellationToken)
    {
        var request = new ProtoWriter();
        request.WriteUInt32(2, ScrProtocol.Program);
        request.WriteString(3, ScrProtocol.GameVersion);
        request.WriteUInt32(4, ScrProtocol.Platform);
        var reply = await CallAsync(GameUtilitiesService, 1, new JsonObject
        {
            ["attribute"] = new JsonArray(
                Attribute("client_request", "classic.protocol.v1.aurora.ConnectToServerRequest"),
                Attribute("protobuf", Convert.ToBase64String(request.ToArray())),
                Attribute("server_instance", "Release")),
        }, cancellationToken).ConfigureAwait(false);

        string? responseType = null, payload = null;
        foreach (var attribute in reply[1]?["attribute"]?.AsArray() ?? [])
        {
            var value = attribute?["value"];
            var text = (value?["blob_value"] ?? value?["string_value"])?.GetValue<string>();
            switch (attribute?["name"]?.GetValue<string>())
            {
                case "client_response": responseType = text; break;
                case "protobuf": payload = text; break;
            }
        }

        if (responseType != "classic.protocol.v1.aurora.ConnectToServerResponse" || payload is null)
        {
            throw new InvalidOperationException($"Battle.net gave an unexpected classic server answer ({responseType ?? "none"}).");
        }

        string? url = null;
        byte[]? ticket = null;
        var r = new ProtoReader(Convert.FromBase64String(payload));
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1 when type == WireType.LengthDelimited: url = r.ReadString(); break;
                case 2 when type == WireType.LengthDelimited: ticket = r.ReadLengthDelimited(); break;
                default: r.Skip(type); break;
            }
        }

        if (url is null || ticket is not { Length: > 0 } || !Uri.TryCreate(url, UriKind.Absolute, out var address)
            || address.Scheme != "wss" || !(address.Host == "battle.net" || address.Host.EndsWith(".battle.net", StringComparison.Ordinal)
                || address.Host.EndsWith(".blizzard.com", StringComparison.Ordinal)))
        {
            throw new InvalidOperationException("Battle.net gave an invalid classic server address.");
        }

        // The client always uses the SC:R RPC path, keeping any query Battle.net added.
        var builder = new UriBuilder(address) { Path = ScrProtocol.ClassicPath };
        return new ClassicEndpoint(builder.Uri, ticket);
    }

    private static JsonObject Attribute(string name, string value) =>
        new() { ["name"] = name, ["value"] = new JsonObject { ["string_value"] = value } };

    private async Task<JsonArray> CallAsync(uint service, uint method, JsonObject body, CancellationToken cancellationToken)
    {
        var reply = await SendAsync(service, method, body, cancellationToken).ConfigureAwait(false);
        return reply;
    }

    private async Task<JsonArray> SendAsync(uint service, uint method, JsonObject body, CancellationToken cancellationToken)
    {
        var waiter = new TaskCompletionSource<JsonArray>(TaskCreationOptions.RunContinuationsAsynchronously);
        ulong token;
        lock (_pending)
        {
            token = _nextToken++;
            _pending[token] = waiter;
        }

        var header = new JsonObject { ["method_id"] = method, ["service_hash"] = service, ["service_id"] = 0, ["token"] = token };
        await WriteAsync(new JsonArray(header, body), cancellationToken).ConfigureAwait(false);
        using (cancellationToken.Register(() => waiter.TrySetCanceled(cancellationToken)))
        {
            var reply = await waiter.Task.ConfigureAwait(false);
            var status = Number(reply[0]?["status"]);
            if (status != 0)
            {
                throw new InvalidOperationException($"Battle.net refused Aurora request {service:X8}/{method} (status {status}).");
            }

            return reply;
        }
    }

    private async Task WriteAsync(JsonArray message, CancellationToken cancellationToken)
    {
        var bytes = Encoding.UTF8.GetBytes(message.ToJsonString());
        await _sendLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            await _socket.SendAsync(bytes, WebSocketMessageType.Text, endOfMessage: true, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _sendLock.Release();
        }
    }

    /// <summary>Reads every message: responses complete their request, Echo is answered, other calls queue for the logon.</summary>
    private async Task ReadLoopAsync()
    {
        var buffer = new byte[16384];
        try
        {
            while (!_stop.IsCancellationRequested)
            {
                using var message = new MemoryStream();
                WebSocketReceiveResult result;
                do
                {
                    result = await _socket.ReceiveAsync(buffer, _stop.Token).ConfigureAwait(false);
                    if (result.MessageType == WebSocketMessageType.Close)
                    {
                        _trace("Battle.net closed the Aurora connection.");
                        return;
                    }

                    message.Write(buffer, 0, result.Count);
                }
                while (!result.EndOfMessage);

                if (JsonNode.Parse(message.ToArray()) is not JsonArray { Count: 2 } pair || pair[0] is not JsonObject header)
                {
                    continue;
                }

                if (IsTrue(header["is_response"]))
                {
                    TaskCompletionSource<JsonArray>? waiter;
                    lock (_pending)
                    {
                        _pending.Remove(Number(header["token"]), out waiter);
                    }

                    waiter?.TrySetResult(pair);
                }
                else if (Number(header["service_hash"]) == ConnectionService && Number(header["method_id"]) == 3)
                {
                    var echo = new JsonObject { ["service_id"] = 254, ["token"] = header["token"]?.DeepClone(), ["is_response"] = true, ["status"] = 0 };
                    await WriteAsync(new JsonArray(echo, pair[1]?.DeepClone()), _stop.Token).ConfigureAwait(false);
                }
                else
                {
                    await _calls.Writer.WriteAsync(pair, _stop.Token).ConfigureAwait(false);
                }
            }
        }
        catch (Exception ex) when (ex is OperationCanceledException or WebSocketException or ObjectDisposedException)
        {
            _trace($"Aurora connection ended: {ex.Message}");
        }
        finally
        {
            lock (_pending)
            {
                foreach (var waiter in _pending.Values)
                {
                    waiter.TrySetException(new InvalidOperationException("The Aurora connection closed."));
                }

                _pending.Clear();
            }

            _calls.Writer.TryComplete(new InvalidOperationException("The Aurora connection closed."));
        }
    }

    private static AuroraSession ParseLogonResult(JsonObject body)
    {
        if (Number(body["error_code"]) is var error and not 0)
        {
            throw new InvalidOperationException($"Battle.net sign-in failed (error {error}).");
        }

        var key = body["session_key"]?.GetValue<string>() ?? throw new InvalidOperationException("Battle.net sent no session key.");
        var account = body["account_id"] ?? throw new InvalidOperationException("Battle.net sent no account ID.");
        var game = body["game_account_id"]?.AsArray().FirstOrDefault() ?? throw new InvalidOperationException("This account has no StarCraft: Remastered game account.");
        return new AuroraSession(
            Convert.FromBase64String(key),
            Number(account["high"]), Number(account["low"]),
            Number(game["high"]), Number(game["low"]),
            Number(body["connected_region"]),
            body["battle_tag"]?.GetValue<string>());
    }

    private static Uri ParseChallengeUrl(JsonObject body)
    {
        if (body["payload_type"]?.GetValue<string>() != "web_auth_url" || body["payload"]?.GetValue<string>() is not { } payload)
        {
            throw new InvalidOperationException("Battle.net asked for an unsupported kind of sign-in.");
        }

        var url = new Uri(Encoding.UTF8.GetString(Convert.FromBase64String(payload)));
        if (url.Scheme != "https" || !(url.Host == "battle.net" || url.Host.EndsWith(".battle.net", StringComparison.Ordinal)))
        {
            throw new InvalidOperationException("Battle.net returned an unexpected sign-in address.");
        }

        return url;
    }

    /// <summary>Aurora sends numbers as JSON numbers or as strings.</summary>
    private static ulong Number(JsonNode? node) => node switch
    {
        JsonValue v when v.TryGetValue<ulong>(out var n) => n,
        JsonValue v when v.TryGetValue<long>(out var n) && n >= 0 => (ulong)n,
        JsonValue v when v.TryGetValue<string>(out var s) && ulong.TryParse(s, out var n) => n,
        _ => 0,
    };

    private static bool IsTrue(JsonNode? node) => node switch
    {
        JsonValue v when v.TryGetValue<bool>(out var b) => b,
        JsonValue v when v.TryGetValue<string>(out var s) => s.Equals("true", StringComparison.OrdinalIgnoreCase),
        _ => false,
    };

    private static uint FourCc(string code) => code.Aggregate(0u, (value, c) => (value << 8) | c);

    public async ValueTask DisposeAsync()
    {
        // Leaving properly: ConnectionService.RequestDisconnect (method 7, no reply), as a Battle.net
        // client does when it logs out. Just dropping the socket leaves the account's SC:R session
        // live for 20-50 seconds, and a new sign-in stalls until it ends.
        if (_socket.State == WebSocketState.Open)
        {
            try
            {
                using var goodbye = new CancellationTokenSource(TimeSpan.FromSeconds(1));
                var header = new JsonObject { ["method_id"] = 7, ["service_hash"] = ConnectionService, ["service_id"] = 0, ["token"] = _nextToken++ };
                await WriteAsync(new JsonArray(header, new JsonObject { ["error_code"] = 0 }), goodbye.Token).ConfigureAwait(false);
                _trace("-> RequestDisconnect");
            }
            catch (Exception ex) when (ex is WebSocketException or OperationCanceledException or ObjectDisposedException)
            {
                // Already going.
            }
        }

        _stop.Cancel();
        try
        {
            if (_socket.State == WebSocketState.Open)
            {
                using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(2));
                await _socket.CloseAsync(WebSocketCloseStatus.NormalClosure, null, timeout.Token).ConfigureAwait(false);
            }
        }
        catch (Exception ex) when (ex is WebSocketException or OperationCanceledException)
        {
            // Already gone.
        }

        if (_reading is not null)
        {
            await _reading.ConfigureAwait(ConfigureAwaitOptions.SuppressThrowing);
        }

        _socket.Dispose();
    }
}
