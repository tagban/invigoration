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
public sealed class ScrConnection : IAsyncDisposable
{
    private static readonly TimeSpan StepTimeout = TimeSpan.FromSeconds(20);

    private readonly AuroraClient _aurora;
    private readonly ClassicWebSocket _classic = new();
    private readonly Action<string> _trace;
    private readonly Channel<ScrChatEvent> _events = Channel.CreateUnbounded<ScrChatEvent>();
    private readonly Dictionary<uint, TaskCompletionSource<ClassicRpc>> _pending = new();
    private readonly Dictionary<ulong, ScrChannel> _catalog = new();
    private readonly CancellationTokenSource _stop = new();
    private TaskCompletionSource _catalogArrived = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private TaskCompletionSource<ScrChannel> _entered = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private ScrChatSession _chat = null!;
    private Task? _reading;
    private Action<byte[]>? _saveCredential;

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

    /// <summary>Signs in and joins <paramref name="homeChannel"/> (a public channel's name). Throws if any step fails.</summary>
    public static async Task<ScrConnection> ConnectAsync(
        byte[]? savedCredential,
        string homeChannel,
        Func<Uri, CancellationToken, Task<byte[]>> challenge,
        Action<string> trace,
        CancellationToken cancellationToken,
        Action<byte[]>? saveCredential = null)
    {
        var connection = new ScrConnection(trace) { _saveCredential = saveCredential };
        try
        {
            await connection.RunSignInAsync(savedCredential, homeChannel, challenge, cancellationToken).ConfigureAwait(false);
            return connection;
        }
        catch
        {
            await connection.DisposeAsync().ConfigureAwait(false);
            throw;
        }
    }

    private async Task RunSignInAsync(byte[]? savedCredential, string homeChannel, Func<Uri, CancellationToken, Task<byte[]>> challenge, CancellationToken cancellationToken)
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
        await CallAsync(ScrProtocol.GameAccountService, ScrProtocol.GetToonsMethod, [], null, cancellationToken).ConfigureAwait(false);
        var version = new ProtoWriter();
        version.WriteString(1, ScrProtocol.GameVersion);
        await CallAsync(ScrProtocol.GameVersionService, ScrProtocol.SetGameVersionMethod, version.ToArray(), null, cancellationToken).ConfigureAwait(false);
        var legacy = new ProtoWriter();
        legacy.WriteUInt32(1, ScrProtocol.LegacyClient);
        await CallAsync(ScrProtocol.LegacyService, ScrProtocol.LegacyConnectMethod, legacy.ToArray(), null, cancellationToken).ConfigureAwait(false);
        _trace("Game session started.");

        await CallAsync(LegacyChatService.Hash, LegacyChatService.SetOnlineMethod, [], null, cancellationToken).ConfigureAwait(false);
        var chatConnect = new ProtoWriter();
        chatConnect.WriteUInt32(1, 0);
        await CallAsync(LegacyChatService.Hash, ScrProtocol.LegacyChatConnectMethod, chatConnect.ToArray(), null, cancellationToken).ConfigureAwait(false);
        await _catalogArrived.Task.WaitAsync(StepTimeout, cancellationToken).ConfigureAwait(false);
        _trace($"Chat online: {Channels.Count} public channel(s).");

        await JoinAsync(homeChannel, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Joins a public channel by its name (case-insensitive) and waits for Battle.net to confirm it.</summary>
    public async Task<ScrChannel> JoinAsync(string channelName, CancellationToken cancellationToken)
    {
        var target = Channels.FirstOrDefault(c => c.DisplayName.Equals(channelName, StringComparison.OrdinalIgnoreCase)
                                                  || c.InternalName.Equals(channelName, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"Battle.net doesn't list a channel called \"{channelName}\".");
        _entered = new TaskCompletionSource<ScrChannel>(TaskCreationOptions.RunContinuationsAsynchronously);
        _trace($"-> join {target.DisplayName} (#{target.Id})");
        await SendAsync(_chat.JoinListedChannel(target), cancellationToken).ConfigureAwait(false);
        var joined = await _entered.Task.WaitAsync(StepTimeout, cancellationToken).ConfigureAwait(false);
        _trace($"Joined {joined.DisplayName}.");
        return joined;
    }

    /// <summary>Sends a message built by <see cref="Chat"/> (a chat line, a whisper, a join).</summary>
    public Task SendAsync(byte[] message, CancellationToken cancellationToken) => _classic.SendAsync(message, true, cancellationToken);

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
                _entered.TrySetResult(entered.Channel);
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
        _stop.Cancel();
        await _classic.DisposeAsync().ConfigureAwait(false);
        if (_reading is not null)
        {
            await _reading.ConfigureAwait(ConfigureAwaitOptions.SuppressThrowing);
        }

        await _aurora.DisposeAsync().ConfigureAwait(false);
    }
}
