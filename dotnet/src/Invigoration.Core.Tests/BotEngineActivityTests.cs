using System.Net;
using System.Net.Sockets;
using Invigoration.Core.Config;
using Invigoration.Core.Networking;
using Invigoration.Core.Protocol;

namespace Invigoration.Core.Tests;

/// <summary>
/// BotEngine.IsIdle decides whether a bot's Connect button shows at all, so getting it wrong either
/// hides the only way to connect a bot that's sitting idle, or offers Connect over the top of one
/// that's still busy. Driven with real local sockets, since every state it reads is a socket's own
/// or set by the engine's handling of one.
/// </summary>
public class BotEngineActivityTests
{
    [Fact]
    public async Task ANewBotIsIdle()
    {
        await using var engine = new BotEngine(new BotConfig());

        Assert.True(engine.IsIdle);
        Assert.False(engine.IsReconnecting);
    }

    [Fact]
    public async Task DebugMode_SaysSoOnlyWhenItActuallyChanges()
    {
        await using var engine = new BotEngine(new BotConfig());
        var changes = 0;
        engine.DebugModeChanged += () => changes++;

        engine.DebugMode = true;
        engine.DebugMode = true;
        engine.DebugMode = false;

        Assert.Equal(2, changes);
    }

    /// <summary>The Bot menu's checkmark has to follow /debug too, not only its own clicks.</summary>
    [Fact]
    public async Task DebugCommand_IsSeenByAnyoneWatchingDebugMode()
    {
        await using var engine = new BotEngine(new BotConfig());
        var changes = 0;
        engine.DebugModeChanged += () => changes++;

        await engine.RunLocalCommandAsync("/debug");

        Assert.True(engine.DebugMode);
        Assert.Equal(1, changes);
    }

    [Fact]
    public async Task AChatConnection_IsBusyWhileUp_AndIdleOnceTheServerHangsUp()
    {
        using var listener = StartListener(out var port);
        await using var engine = new BotEngine(ChatBot(port));
        var notifications = 0;
        engine.ActivityChanged += () => Interlocked.Increment(ref notifications);

        await engine.ConnectAsync();
        using var server = await listener.AcceptTcpClientAsync().WaitAsync(TimeSpan.FromSeconds(5));

        Assert.True(await Waited(() => !engine.IsIdle), "a live connection should keep the bot busy");
        Assert.True(Volatile.Read(ref notifications) > 0);

        server.Close();

        Assert.True(await Waited(() => engine.IsIdle), "with auto-reconnect off, a dropped bot has nothing left to do");
        Assert.False(engine.IsReconnecting);
    }

    /// <summary>
    /// A reconnect countdown has no socket up at all. It isn't idle — Disconnect has to be able to
    /// stop it — but it is only waiting, so Connect stays on offer to skip the wait.
    /// </summary>
    [Fact]
    public async Task AReconnectCountdown_IsOnlyWaiting_UntilDisconnectStopsIt()
    {
        using var listener = StartListener(out var port);
        var config = ChatBot(port);
        config.AutoReconnect = true;
        config.AutoReconnectDelaySeconds = 60;
        await using var engine = new BotEngine(config);

        await engine.ConnectAsync();
        using (await listener.AcceptTcpClientAsync().WaitAsync(TimeSpan.FromSeconds(5)))
        {
        }

        Assert.True(await Waited(() => engine.IsReconnecting), "the drop should have started a reconnect countdown");
        Assert.False(engine.IsIdle);
        Assert.True(engine.IsWaitingToReconnect);

        await engine.DisconnectAsync();

        Assert.True(await Waited(() => engine.IsIdle), "Disconnect should stop the countdown");
        Assert.False(engine.IsReconnecting);
        Assert.False(engine.IsWaitingToReconnect);
    }

    /// <summary>While an attempt is actually in flight it isn't just waiting — Connect then would start a second logon over the first.</summary>
    [Fact]
    public async Task AConnectInFlight_IsNotJustWaiting()
    {
        using var listener = StartListener(out var port);
        await using var engine = new BotEngine(ClassicBot(bnlsPort: port));

        await engine.ConnectAsync();
        using var bnls = await listener.AcceptTcpClientAsync().WaitAsync(TimeSpan.FromSeconds(5));

        Assert.False(engine.IsIdle);
        Assert.False(engine.IsWaitingToReconnect);
    }

    /// <summary>
    /// Between BNLS answering and Battle.net itself connecting, BNLS is the only socket up — and
    /// BNLS alone can't count, since it can stay open after logon too. The pending logon covers
    /// that stretch, and ends when BNLS goes away without getting it any further.
    /// </summary>
    [Fact]
    public async Task AClassicLogon_IsBusyWhileOnlyBnlsIsUp_AndIdleOnceBnlsHangsUp()
    {
        using var listener = StartListener(out var port);
        await using var engine = new BotEngine(ClassicBot(bnlsPort: port));

        await engine.ConnectAsync();
        using var bnls = await listener.AcceptTcpClientAsync().WaitAsync(TimeSpan.FromSeconds(5));

        Assert.False(engine.IsIdle);

        bnls.Close();

        Assert.True(await Waited(() => engine.IsIdle), "a logon BNLS walked away from can't get any further");
    }

    [Fact]
    public async Task Disconnect_EndsALogonStuckWaitingOnBnls()
    {
        using var listener = StartListener(out var port);
        await using var engine = new BotEngine(ClassicBot(bnlsPort: port));

        await engine.ConnectAsync();
        using var silentBnls = await listener.AcceptTcpClientAsync().WaitAsync(TimeSpan.FromSeconds(5));
        Assert.False(engine.IsIdle);

        await engine.DisconnectAsync();

        Assert.True(engine.IsIdle);
    }

    /// <summary>
    /// A connect can sit waiting for as long as the OS lets it with no socket up to show for it —
    /// here a SOCKS proxy that accepts and then never answers. That counts as busy, and Disconnect
    /// stops it rather than letting the bot come up afterwards anyway — quietly: ConnectAsync
    /// returns rather than throwing, since callers like Normalize Password and /reconnect have
    /// nowhere to catch a cancellation the user chose.
    /// </summary>
    [Fact]
    public async Task Disconnect_CancelsAConnectStillWaitingOnItsSocket()
    {
        using var proxy = StartListener(out var proxyPort);
        var config = ClassicBot(bnlsPort: 9367);
        config.ProxyEnabled = true;
        config.ProxyProtocol = ProxyProtocol.Socks5;
        config.ProxyHost = "127.0.0.1";
        config.ProxyPort = proxyPort;
        await using var engine = new BotEngine(config);

        var connecting = engine.ConnectAsync();
        using var silentProxy = await proxy.AcceptTcpClientAsync().WaitAsync(TimeSpan.FromSeconds(5));
        Assert.False(engine.IsIdle);
        Assert.False(connecting.IsCompleted);

        await engine.DisconnectAsync();

        await connecting.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.True(await Waited(() => engine.IsIdle));
    }

    /// <summary>
    /// A removed bot used to keep any reconnect countdown it had going — nothing stopped it — and
    /// come back online with no tab left to show it.
    /// </summary>
    [Fact]
    public async Task RemovingABot_StopsItsReconnectCountdown()
    {
        using var listener = StartListener(out var port);
        var config = ChatBot(port);
        config.AutoReconnect = true;
        config.AutoReconnectDelaySeconds = 3;
        var engine = new BotEngine(config);

        await engine.ConnectAsync();
        using (await listener.AcceptTcpClientAsync().WaitAsync(TimeSpan.FromSeconds(5)))
        {
        }

        Assert.True(await Waited(() => engine.IsReconnecting), "the drop should have started a reconnect countdown");

        await engine.DisposeAsync();

        // Checked well inside the 3s delay, so a slow machine can't turn the old bug into a pass
        // or the fix into a failure; then past it, to be sure no attempt went out after all.
        Assert.True(await Waited(() => !engine.IsReconnecting, 1500), "removing the bot should stop its countdown");
        await Task.Delay(2500);
        Assert.False(listener.Pending(), "a removed bot reconnected");
    }

    private static TcpListener StartListener(out int port)
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        port = ((IPEndPoint)listener.LocalEndpoint).Port;
        return listener;
    }

    private static BotConfig ChatBot(int port) => new()
    {
        Product = BncsProduct.Chat,
        Username = "SomeUser",
        Password = "somepassword",
        BattlenetServer = "127.0.0.1",
        BattlenetPort = port,
    };

    private static BotConfig ClassicBot(int bnlsPort) => new()
    {
        Product = BncsProduct.Diablo,
        Username = "SomeUser",
        Password = "somepassword",
        BattlenetServer = "127.0.0.1",
        BattlenetPort = 1,
        BnlsServer = "127.0.0.1",
        BnlsPort = bnlsPort,
    };

    /// <summary>Polls until the condition holds or the timeout passes — for state a background receive loop sets.</summary>
    private static async Task<bool> Waited(Func<bool> condition, int timeoutMs = 5000)
    {
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
        while (DateTime.UtcNow < deadline)
        {
            if (condition())
            {
                return true;
            }

            await Task.Delay(25);
        }

        return condition();
    }
}
