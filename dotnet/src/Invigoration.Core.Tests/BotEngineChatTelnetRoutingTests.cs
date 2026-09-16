using System.Net;
using System.Net.Sockets;
using System.Reflection;
using Invigoration.Core.Auth;
using Invigoration.Core.Config;
using Invigoration.Core.Protocol;

namespace Invigoration.Core.Tests;

/// <summary>
/// Chat / Telnet is offered in the Game list like a product, but it's a different connection type:
/// no BNLS hop, no CD key, no version check. These cover the branches that decide that — picking it
/// must route the connect straight at the Battle.net server, not at BNLS.
/// In the LogonCheckCache collection because CanLogOnWithoutBnls reads that process-wide cache (and
/// this clears it), so these must not run alongside the rapid-reconnect tests that populate it.
/// </summary>
[Collection("LogonCheckCache")]
public class BotEngineChatTelnetRoutingTests
{
    private static BotConfig ChatConfig() => new()
    {
        Product = BncsProduct.Chat,
        Username = "Tagban",
        Password = "huntertwo",
        // Nothing listens on port 1 — the connect fails either way; which host it was aimed at is the point.
        BattlenetServer = "127.0.0.1",
        BattlenetPort = 1,
        BnlsServer = "127.0.0.1",
        BnlsPort = 1,
    };

    [Fact]
    public void ChatTelnet_NeedsNoCdKeyAndIsPrivateServersOnly()
    {
        Assert.True(BncsProduct.IsChatTelnet(BncsProduct.Chat));
        Assert.False(BncsProduct.RequiresCdKey(BncsProduct.Chat));
        Assert.False(BncsProduct.RequiresExpansionCdKey(BncsProduct.Chat));
        Assert.Equal(ServerCompatibility.PrivateOnly, BncsProduct.GetServerCompatibility(BncsProduct.Chat));
        Assert.Equal("Chat / Telnet", BncsProduct.GetDisplayName(BncsProduct.Chat));

        // Kept out of the catalog of real BNCS products on purpose — it isn't one.
        Assert.DoesNotContain(BncsProduct.Chat, BncsProduct.Catalog.Keys);
    }

    [Fact]
    public async Task Connecting_GoesStraightToTheServerInsteadOfBnls()
    {
        await using var engine = new BotEngine(ChatConfig());
        var logged = new List<string>();
        engine.Log += segments => { lock (logged) { logged.Add(string.Concat(segments.Select(s => s.Text))); } };

        try
        {
            await engine.ConnectAsync();
        }
        catch (SocketException)
        {
            // Nothing is listening; only which connection the engine chose matters here.
        }

        Assert.Contains(logged, l => l.Contains("(Chat protocol)", StringComparison.Ordinal));
        Assert.DoesNotContain(logged, l => l.Contains("Login Server connecting", StringComparison.Ordinal));
    }

    /// <summary>
    /// A normal game can only skip BNLS once it has cached answers from an earlier logon this run;
    /// the Chat protocol never asks BNLS anything, so it qualifies with nothing cached at all.
    /// </summary>
    [Fact]
    public async Task ChatTelnet_CanAlwaysLogOnWithoutBnls()
    {
        LogonCheckCache.ClearForTests();

        await using var chat = new BotEngine(ChatConfig());
        await using var game = new BotEngine(new BotConfig { Product = BncsProduct.Warcraft2BNE, CdKey = "" });

        Assert.True(chat.CanLogOnWithoutBnls());
        Assert.False(game.CanLogOnWithoutBnls());
    }

    /// <summary>Rapid reconnect still only applies off official Battle.net — the Chat protocol doesn't change that.</summary>
    [Theory]
    [InlineData("war.pianka.io", true)]
    [InlineData("useast.battle.net", false)]
    public async Task RapidReconnect_AppliesOnPrivateServersOnly(string server, bool expected)
    {
        var config = ChatConfig();
        config.BattlenetServer = server;
        config.RapidReconnect = true;
        await using var engine = new BotEngine(config);

        Assert.Equal(expected, engine.RapidReconnectApplies);
    }

    /// <summary>
    /// Chat has no SID_CHATCOMMAND — the line goes out as typed. Asserted through the send gate's
    /// own refusal message, which names the state that blocked it, so this stays a check of the
    /// routing decision rather than needing a live socket.
    /// </summary>
    [Fact]
    public async Task SendingBeforeLogon_IsRefusedTheSameWayAsBncs()
    {
        await using var engine = new BotEngine(ChatConfig());
        var logged = new List<string>();
        engine.Log += segments => { lock (logged) { logged.Add(string.Concat(segments.Select(s => s.Text))); } };

        await engine.SendChatCommandAsync("hello");

        Assert.Contains(logged, l => l.Contains("not logged on to Battle.net", StringComparison.Ordinal));
    }

    /// <summary>Once the Chat login completes, the shared logged-on flag is what the rest of the engine reads — chat sends, uptime, and the reconnect loop's exit condition all key off it.</summary>
    [Fact]
    public async Task LoggingOn_MarksTheSessionOnForTheRestOfTheEngine()
    {
        await using var engine = new BotEngine(ChatConfig());
        var handle = typeof(BotEngine).GetMethod("HandleChatTelnetLineAsync", BindingFlags.NonPublic | BindingFlags.Instance)!;

        await (Task)handle.Invoke(engine, ["2010 NAME Tagban"])!;

        var auth = typeof(BotEngine).GetField("_auth", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(engine)!;
        Assert.True((bool)auth.GetType().GetProperty("LoggedOnToBncs")!.GetValue(auth)!);
    }

    /// <summary>
    /// war.pianka.io (probed live 2026-09-15) accepts the connection and then says nothing at all —
    /// no banner, no "Username:" — until it's given a username and password. A bot that waits for a
    /// prompt there waits forever, which is the state this covers: connect to a server that stays
    /// silent and the login must still go out on its own, handshake first.
    /// </summary>
    [Fact]
    public async Task ASilentServer_StillGetsTheLoginWithoutEverPrompting()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;

        var received = new List<byte>();
        var accepted = Task.Run(async () =>
        {
            using var client = await listener.AcceptTcpClientAsync();
            using var stream = client.GetStream();
            var buffer = new byte[256];
            var deadline = DateTime.UtcNow.AddSeconds(5);
            // Read until both credential lines have arrived, or time's up — never replying, which
            // is the whole point of the scenario.
            while (DateTime.UtcNow < deadline && received.Count(b => b == (byte)'\n') < 2)
            {
                var read = await stream.ReadAsync(buffer);
                if (read == 0)
                {
                    break;
                }

                received.AddRange(buffer[..read]);
            }
        });

        var config = ChatConfig();
        config.BattlenetServer = "127.0.0.1";
        config.BattlenetPort = port;
        await using var engine = new BotEngine(config);
        await engine.ConnectAsync();
        await accepted.WaitAsync(TimeSpan.FromSeconds(10));

        // Handshake selector bytes first, then the two credential lines.
        Assert.Equal(0x03, received[0]);
        Assert.Equal(0x04, received[1]);
        var lines = System.Text.Encoding.UTF8.GetString(received.Skip(2).ToArray())
            .Split("\r\n", StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(["Tagban", "huntertwo"], lines);
    }
}
