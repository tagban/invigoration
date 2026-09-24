using System.Threading.Channels;
using Invigoration.Sc2.Protobuf;
using Invigoration.Scr.Aurora;
using Invigoration.Scr.Classic;
using Invigoration.Scr.LegacyChat;

namespace Invigoration.Scr;

/// <summary>
/// A native StarCraft: Remastered chat connection, start to finish: Aurora logon (with the web
/// sign-in only if Battle.net asks), the classic server's address and ticket, the classic
/// WebSocket with its key-folded seed, AuthSession, the startup calls, and the home channel.
/// Then it keeps both sockets answered and hands chat events to the caller. The sequence follows
/// ncarrillo/sc1-research (MIT).
/// </summary>
/// <summary>Which character to play as and where to start. The character is looked up on the gateway.</summary>
public sealed record ScrConnectOptions(uint Gateway = ScrGateways.UsEast, string? CharacterName = null, string HomeChannel = ScrConnectOptions.DefaultHomeChannel)
{
    public const string DefaultHomeChannel = "Open Tech Support";
}

public sealed class ScrConnection : IAsyncDisposable
{
    private static readonly TimeSpan StepTimeout = TimeSpan.FromSeconds(20);

    private readonly AuroraClient _aurora;
    private readonly ClassicWebSocket _classic = new();
    private readonly Action<string> _trace;
    private readonly Channel<ScrChatEvent> _events = Channel.CreateUnbounded<ScrChatEvent>();
    private readonly Dictionary<uint, TaskCompletionSource<ClassicRpc>> _pending = new();
    private readonly Dictionary<ulong, ScrChannel> _catalog = new();
    private readonly List<ScrGateway> _gateways = new();
    private readonly CancellationTokenSource _stop = new();
    private TaskCompletionSource _catalogArrived = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private TaskCompletionSource<ScrChannel> _entered = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private ScrChatSession _chat = null!;
    private Task? _reading;
    private Action<byte[]>? _saveCredential;
    private bool _chatOnline;

    private ScrConnection(Action<string> trace)
    {
        _trace = trace;
        _aurora = new AuroraClient(trace);
    }

    public ScrChatSession Chat => _chat;

    public string? BattleTag { get; private set; }

    /// <summary>A fresh "keep me signed in" credential issued during this sign-in, to save for next time. Null if Battle.net didn't issue one.</summary>
    public byte[]? NewSavedCredential { get; private set; }

    /// <summary>Every chat event, from the channel list during startup onwards. Completes when the classic connection closes.</summary>
    public ChannelReader<ScrChatEvent> Events => _events.Reader;

    /// <summary>The public channels Battle.net advertised, by ID.</summary>
    public IReadOnlyCollection<ScrChannel> Channels
    {
        get
        {
            lock (_catalog)
            {
                return _catalog.Values.ToList();
            }
        }
    }

    /// <summary>The characters Battle.net listed for this account.</summary>
    public IReadOnlyList<ScrToon> Toons { get; private set; } = [];

    /// <summary>The character this connection plays as, once chosen.</summary>
    public ScrToon? Toon { get; private set; }

    /// <summary>The gateways Battle.net announced at sign-in, in the order it sent them.</summary>
    public IReadOnlyList<ScrGateway> Gateways
    {
        get
        {
            lock (_gateways)
            {
                return _gateways.ToList();
            }
        }
    }

    /// <summary>
    /// Signs in, picks the character on <see cref="ScrConnectOptions.Gateway"/>, and joins the home
    /// channel. Throws if any step fails; <see cref="ScrCharacterException"/> when the account has no
    /// character there.
    /// </summary>
    public static async Task<ScrConnection> ConnectAsync(
        byte[]? savedCredential,
        ScrConnectOptions options,
        Func<Uri, CancellationToken, Task<byte[]>> challenge,
        Action<string> trace,
        CancellationToken cancellationToken,
        Action<byte[]>? saveCredential = null)
    {
        var connection = new ScrConnection(trace) { _saveCredential = saveCredential };
        try
        {
            await connection.SignInAsync(savedCredential, challenge, cancellationToken).ConfigureAwait(false);
            await connection.StartGameAsync(options, cancellationToken).ConfigureAwait(false);
            return connection;
        }
        catch
        {
            await connection.DisposeAsync().ConfigureAwait(false);
            throw;
        }
    }

    /// <summary>
    /// Signs in only as far as the character list, for choosing a character and gateway before
    /// connecting, then closes. Uses (and renews) the saved sign-in like a real connect.
    /// </summary>
    public static async Task<ScrAccount> ListCharactersAsync(
        byte[]? savedCredential,
        Func<Uri, CancellationToken, Task<byte[]>> challenge,
        Action<string> trace,
        CancellationToken cancellationToken,
        Action<byte[]>? saveCredential = null,
        Func<Func<string, CancellationToken, Task<byte[]?>>, Task>? whileSignedIn = null)
    {
        await using var connection = new ScrConnection(trace) { _saveCredential = saveCredential };
        await connection.SignInAsync(savedCredential, challenge, cancellationToken).ConfigureAwait(false);

        await connection.LoadToonsAsync(cancellationToken).ConfigureAwait(false);

        // After the character list: Battle.net wants that within seconds of the classic socket opening.
        if (whileSignedIn is not null)
        {
            await whileSignedIn(connection.GenerateWebCredentialsAsync).ConfigureAwait(false);
        }

        var gateways = connection.Gateways;
        return new ScrAccount(connection.BattleTag, connection.Toons, gateways.Count > 0 ? gateways : ScrGateways.Known);
    }

    /// <summary>
    /// Creates a character on <paramref name="gateway"/> with GameAccount.CreateToon, before any game
    /// session starts (as the retail client's character screen does), then lists the characters
    /// again. Returns the character Battle.net reports creating, or null, and the list afterwards.
    /// </summary>
    public static async Task<(ScrToon? Created, ScrAccount After)> CreateToonAsync(
        byte[]? savedCredential,
        string name,
        uint gateway,
        Func<Uri, CancellationToken, Task<byte[]>> challenge,
        Action<string> trace,
        CancellationToken cancellationToken,
        Action<byte[]>? saveCredential = null)
    {
        await using var connection = new ScrConnection(trace) { _saveCredential = saveCredential };
        await connection.SignInAsync(savedCredential, challenge, cancellationToken).ConfigureAwait(false);
        await connection.LoadToonsAsync(cancellationToken).ConfigureAwait(false);

        var request = new ProtoWriter();
        request.WriteString(1, name);
        request.WriteUInt64(2, gateway);
        var reply = await connection.CallAsync(ScrProtocol.GameAccountService, ScrProtocol.CreateToonMethod, request.ToArray(), null, cancellationToken).ConfigureAwait(false);
        trace($"CreateToon reply: header {reply.Header}, body {Convert.ToHexString(reply.Body)}");
        var created = ScrToons.Decode(reply.Body).FirstOrDefault();

        await connection.LoadToonsAsync(cancellationToken).ConfigureAwait(false);
        var gateways = connection.Gateways;
        return (created, new ScrAccount(connection.BattleTag, connection.Toons, gateways.Count > 0 ? gateways : ScrGateways.Known));
    }

    /// <summary>Aurora logon, the classic server's address and ticket, the classic socket, and AuthSession.</summary>
    private async Task SignInAsync(byte[]? savedCredential, Func<Uri, CancellationToken, Task<byte[]>> challenge, CancellationToken cancellationToken)
    {
        _trace($"Opening Aurora at {AuroraClient.UsEndpoint}");
        await _aurora.ConnectAsync(AuroraClient.UsEndpoint, cancellationToken).ConfigureAwait(false);
        _trace(savedCredential is null ? "No saved sign-in; Battle.net will ask for one." : "Presenting the saved sign-in.");
        var session = await _aurora.LogonAsync(ScrProtocol.ProgramCode, savedCredential, challenge, cancellationToken).ConfigureAwait(false);
        BattleTag = session.BattleTag;

        try
        {
            NewSavedCredential = await _aurora.GenerateWebCredentialsAsync(ScrProtocol.ProgramCode, cancellationToken).ConfigureAwait(false);
            _trace(NewSavedCredential is null ? "Battle.net issued no reusable sign-in." : "Battle.net issued a new reusable sign-in.");
            if (NewSavedCredential is { } fresh)
            {
                // Saved now, not at the end: the sign-in is valid even if a later step fails.
                _saveCredential?.Invoke(fresh);
            }
        }
        catch (InvalidOperationException ex)
        {
            _trace($"Couldn't get a reusable sign-in: {ex.Message}");
        }

        var endpoint = await _aurora.ConnectToServerAsync(cancellationToken).ConfigureAwait(false);
        _trace($"Classic server: {endpoint.Url.Host} ({endpoint.Ticket.Length}-byte ticket).");

        await _classic.ConnectAsync(endpoint.Url, [], cancellationToken).ConfigureAwait(false);
        var seed = ClassicEnvelope.SeedFromWebSocketKey(_classic.Key);
        _chat = new ScrChatSession(seed, token: 0, channelId: 0);
        _reading = Task.Run(ReadClassicAsync);
        _trace("Classic WebSocket open. Sending AuthSession...");

        var reply = await CallAsync(ScrProtocol.AuthenticationService, ScrProtocol.AuthSessionMethod, AuthSessionBody(endpoint, session), ScrProtocol.NewRequestTrace(), cancellationToken).ConfigureAwait(false);
        var proof = ProofLength(reply.Body);
        if (proof is not (48 or 64))
        {
            throw new InvalidOperationException($"Battle.net's AuthSession reply has no valid server proof ({proof} bytes).");
        }

        _trace($"AuthSession accepted ({proof}-byte server proof).");
    }

    private async Task LoadToonsAsync(CancellationToken cancellationToken)
    {
        var toons = await CallAsync(ScrProtocol.GameAccountService, ScrProtocol.GetToonsMethod, [], null, cancellationToken).ConfigureAwait(false);
        Toons = ScrToons.Decode(toons.Body);
        _trace(Toons.Count == 0 ? "The account has no characters." : $"Characters: {string.Join(", ", Toons)}.");
    }

    /// <summary>
    /// The rest of the startup, which Battle.net wants within about six seconds of the socket
    /// opening: the character list, the game version, Legacy.Connect with the chosen character,
    /// then LegacyChat and the home channel.
    /// </summary>
    private async Task StartGameAsync(ScrConnectOptions options, CancellationToken cancellationToken)
    {
        await LoadToonsAsync(cancellationToken).ConfigureAwait(false);
        Toon = ScrToons.Choose(Toons, options.Gateway, options.CharacterName);
        _trace($"Playing as {Toon}.");

        var version = new ProtoWriter();
        version.WriteString(1, ScrProtocol.GameVersion);
        await CallAsync(ScrProtocol.GameVersionService, ScrProtocol.SetGameVersionMethod, version.ToArray(), null, cancellationToken).ConfigureAwait(false);

        // Legacy.Connect's field 1 is the character ID from GetToons: the retail client's capture
        // sent 3 for the character with ID 3, and the gateway follows from the character.
        var legacy = new ProtoWriter();
        legacy.WriteUInt32(1, Toon.Id);
        await CallAsync(ScrProtocol.LegacyService, ScrProtocol.LegacyConnectMethod, legacy.ToArray(), null, cancellationToken).ConfigureAwait(false);
        _trace("Game session started.");

        await CallAsync(LegacyChatService.Hash, LegacyChatService.SetOnlineMethod, [], null, cancellationToken).ConfigureAwait(false);
        var chatConnect = new ProtoWriter();
        chatConnect.WriteUInt32(1, 0);
        await CallAsync(LegacyChatService.Hash, ScrProtocol.LegacyChatConnectMethod, chatConnect.ToArray(), null, cancellationToken).ConfigureAwait(false);
        _chatOnline = true;
        await _catalogArrived.Task.WaitAsync(StepTimeout, cancellationToken).ConfigureAwait(false);
        _trace($"Chat online: {Channels.Count} public channel(s).");

        await JoinAsync(options.HomeChannel, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>The channel Battle.net last confirmed we're in, if any.</summary>
    public ScrChannel? CurrentChannel { get; private set; }

    /// <summary>
    /// Joins a channel by name (case-insensitive) and waits for Battle.net to confirm it. Outside
    /// any channel (right after sign-in) a listed channel is joined by its ID, as the game does;
    /// from inside one, Battle.net ignores that, so it's the chat command ("/channel name"), which
    /// works for any channel and creates one that doesn't exist.
    /// </summary>
    public async Task<ScrChannel> JoinAsync(string channelName, CancellationToken cancellationToken)
    {
        if (CurrentChannel is { } current && IsNamed(current, channelName))
        {
            _trace($"Already in {current.DisplayName}.");
            return current;
        }

        var listed = Channels.FirstOrDefault(c => IsNamed(c, channelName));
        if (CurrentChannel is null && listed is null && channelName != ScrConnectOptions.DefaultHomeChannel && Channels.Count > 0)
        {
            // The command runs from a channel: outside one, Battle.net ignores it.
            _trace($"Not in a channel yet; entering {ScrConnectOptions.DefaultHomeChannel} before joining {channelName}.");
            await JoinAsync(ScrConnectOptions.DefaultHomeChannel, cancellationToken).ConfigureAwait(false);
        }

        _entered = new TaskCompletionSource<ScrChannel>(TaskCreationOptions.RunContinuationsAsynchronously);
        if (CurrentChannel is null && listed is not null)
        {
            _trace($"-> join {listed.DisplayName} (#{listed.Id})");
            await SendAsync(_chat.JoinListedChannel(listed), cancellationToken).ConfigureAwait(false);
        }
        else
        {
            _trace($"-> join {channelName} (channel command)");
            await SendAsync(_chat.JoinByName(channelName), cancellationToken).ConfigureAwait(false);
        }

        var joined = await _entered.Task.WaitAsync(StepTimeout, cancellationToken).ConfigureAwait(false);
        _trace($"Joined {joined.DisplayName}.");
        return joined;
    }

    private static bool IsNamed(ScrChannel channel, string name) =>
        channel.DisplayName.Equals(name, StringComparison.OrdinalIgnoreCase)
        || channel.InternalName.Equals(name, StringComparison.OrdinalIgnoreCase);

    /// <summary>A "keep me signed in" credential for <paramref name="program"/>, from this signed-in session. Null if Battle.net issues none.</summary>
    public Task<byte[]?> GenerateWebCredentialsAsync(string program, CancellationToken cancellationToken) =>
        _aurora.GenerateWebCredentialsAsync(program, cancellationToken);

    /// <summary>Sends a message built by <see cref="Chat"/> (a chat line, a whisper, a join).</summary>
    public Task SendAsync(byte[] message, CancellationToken cancellationToken) => _classic.SendAsync(message, true, cancellationToken);

    /// <summary>A call to any classic service (Battle.net whispers, friends...); returns the reply's body.</summary>
    public async Task<byte[]> RequestAsync(uint service, uint method, byte[] body, CancellationToken cancellationToken) =>
        (await CallAsync(service, method, body, null, cancellationToken).ConfigureAwait(false)).Body;

    private async Task<ClassicRpc> CallAsync(uint service, uint method, byte[] body, byte[]? trace, CancellationToken cancellationToken)
    {
        var (message, token) = _chat.Call(service, method, body, trace);
        var waiter = new TaskCompletionSource<ClassicRpc>(TaskCreationOptions.RunContinuationsAsynchronously);
        lock (_pending)
        {
            _pending[token] = waiter;
        }

        _trace($"-> call {service:X8}/{method:X8} #{token} ({body.Length} bytes)");
        await SendAsync(message, cancellationToken).ConfigureAwait(false);
        var reply = await waiter.Task.WaitAsync(StepTimeout, cancellationToken).ConfigureAwait(false);
        _trace($"<- reply #{token} ({reply.Body.Length} bytes)");
        return reply;
    }

    private async Task ReadClassicAsync()
    {
        try
        {
            while (await _classic.ReceiveAsync(_stop.Token).ConfigureAwait(false) is { } message)
            {
                if (!message.Binary)
                {
                    continue;
                }

                var (events, replies) = _chat.Receive(message.Data, out var responses);
                if (replies.Count > 0)
                {
                    _trace($"   answered {replies.Count} server call(s)");
                }

                foreach (var reply in replies)
                {
                    await _classic.SendAsync(reply, true, _stop.Token).ConfigureAwait(false);
                }

                foreach (var response in responses)
                {
                    TaskCompletionSource<ClassicRpc>? waiter;
                    lock (_pending)
                    {
                        _pending.Remove(response.Header.Token, out waiter);
                    }

                    waiter?.TrySetResult(response);
                }

                foreach (var chatEvent in events)
                {
                    Track(chatEvent);
                    await _events.Writer.WriteAsync(chatEvent, _stop.Token).ConfigureAwait(false);
                }
            }

            _trace(_classic.CloseCode is { } code
                ? $"Battle.net closed the classic connection (code {code}, \"{_classic.CloseReason}\")."
                : "Battle.net closed the classic connection.");
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            _trace($"Classic connection failed: {ex.Message}");
        }
        catch (OperationCanceledException)
        {
            // Closed on purpose.
        }
        finally
        {
            lock (_pending)
            {
                foreach (var waiter in _pending.Values)
                {
                    waiter.TrySetException(new InvalidOperationException("The classic connection closed."));
                }

                _pending.Clear();
            }

            _entered.TrySetException(new InvalidOperationException("The classic connection closed."));
            _catalogArrived.TrySetException(new InvalidOperationException("The classic connection closed."));
            _events.Writer.TryComplete();
        }
    }

    private void Track(ScrChatEvent chatEvent)
    {
        _trace(chatEvent switch
        {
            ScrChatEvent.ChannelListChanged l => "   channel list: " + string.Join(", ", l.Changes.Select(c => $"{(c.IsRemoval ? "-" : "+")}{c.Channel.DisplayName}#{c.Channel.Id}({c.Channel.Members.Count})")),
            ScrChatEvent.ChannelEntered c => $"   entered {c.Channel.DisplayName}#{c.Channel.Id} ({c.Channel.Members.Count} members)",
            ScrChatEvent.Unhandled u => $"   call {u.Service:X8}/{u.Method:X8}",
            _ => $"   {chatEvent.GetType().Name}",
        });
        switch (chatEvent)
        {
            case ScrChatEvent.ChannelListChanged list:
                lock (_catalog)
                {
                    foreach (var change in list.Changes)
                    {
                        if (change.IsRemoval)
                        {
                            _catalog.Remove(change.Channel.Id);
                        }
                        else
                        {
                            _catalog[change.Channel.Id] = change.Channel;
                        }
                    }
                }

                _catalogArrived.TrySetResult();
                break;
            case ScrChatEvent.ChannelEntered entered:
                CurrentChannel = entered.Channel;
                _entered.TrySetResult(entered.Channel);
                break;
            case ScrChatEvent.ChannelLeft left when CurrentChannel is { } current && (left.ChannelId == current.Id || left.ChannelId == 0):
                CurrentChannel = null;
                break;
            case ScrChatEvent.Unhandled { Service: ScrProtocol.GatewayService, Method: ScrProtocol.GatewayUpdateMethod } update:
                if (ScrGateways.DecodeUpdate(update.Body) is { } gateway)
                {
                    lock (_gateways)
                    {
                        _gateways.RemoveAll(g => g.Id == gateway.Id);
                        _gateways.Add(gateway);
                    }
                }

                break;
        }
    }

    /// <summary>AuthSession: the ticket, the Aurora session key and IDs, and the retail client's identity.</summary>
    private static byte[] AuthSessionBody(ClassicEndpoint endpoint, AuroraSession session)
    {
        var accounts = new ProtoWriter();
        accounts.WriteUInt64(1, session.AccountHigh);
        accounts.WriteUInt64(2, session.AccountLow);
        accounts.WriteUInt64(3, session.GameAccountHigh);
        accounts.WriteUInt64(4, session.GameAccountLow);

        var info = new ProtoWriter();
        info.WriteBytesField(1, session.SessionKey);
        info.WriteUInt32(2, ScrProtocol.ApplicationVersion);
        info.WriteUInt32(3, ScrProtocol.Locale);
        info.WriteUInt32(4, ScrProtocol.Platform);
        info.WriteBytesField(6, accounts.ToArray());
        info.WriteUInt32(7, ScrProtocol.SessionType);

        var body = new ProtoWriter();
        body.WriteBytesField(1, endpoint.Ticket);
        body.WriteBytesField(2, info.ToArray());
        body.WriteUInt32(3, ScrProtocol.Program);
        body.WriteString(4, ScrProtocol.GameVersion);
        body.WriteString(5, ScrProtocol.ClientIdentity);
        body.WriteUInt32(6, ScrProtocol.ClientCapabilities);
        body.WriteUInt32(7, ScrProtocol.Platform);
        body.WriteUInt32(8, 1);
        return body.ToArray();
    }

    private static int ProofLength(byte[] body)
    {
        var r = new ProtoReader(body);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            if (field == 3 && type == WireType.LengthDelimited)
            {
                return r.ReadLengthDelimited().Length;
            }

            r.Skip(type);
        }

        return 0;
    }

    public async ValueTask DisposeAsync()
    {
        // Say goodbye to LegacyChat first, in case that lets Battle.net end the session sooner:
        // a new sign-in stalls while it still thinks the last one is live.
        if (_chatOnline && _reading is { IsCompleted: false })
        {
            try
            {
                var (goodbye, _) = _chat.Call(LegacyChatService.Hash, ScrProtocol.LegacyChatDisconnectMethod, []);
                using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(1));
                await _classic.SendAsync(goodbye, true, timeout.Token).ConfigureAwait(false);
                await Task.Delay(200, CancellationToken.None).ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                // Already gone.
            }
        }

        _stop.Cancel();
        await _classic.DisposeAsync().ConfigureAwait(false);
        if (_reading is not null)
        {
            await _reading.ConfigureAwait(ConfigureAwaitOptions.SuppressThrowing);
        }

        await _aurora.DisposeAsync().ConfigureAwait(false);
    }
}
