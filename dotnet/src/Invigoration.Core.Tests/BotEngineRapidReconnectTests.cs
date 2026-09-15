using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using Invigoration.Core.Auth;
using Invigoration.Core.Config;
using Invigoration.Core.Networking;
using Invigoration.Core.Protocol;

namespace Invigoration.Core.Tests;

/// <summary>
/// Rapid reconnect (BotEngine.RapidReconnect.cs) against loopback stand-ins for the Battle.net
/// server and BNLS: a cached attempt logs on without touching BNLS, a changed challenge falls back
/// to BNLS, BNLS's answers get remembered, and attempts repeat every second until stopped.
/// LogonCheckCache is shared by the whole process, so these run one at a time.
/// </summary>
[Collection("LogonCheckCache")]
public class BotEngineRapidReconnectTests : IDisposable
{
    private const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Instance;
    private static readonly FileTimeValue ChallengeTime = new(0x11111111, 0x01D00000);
    private const string ChallengeFile = "ver-IX86-3.mpq";
    private const string ChallengeFormula = "A=1 B=2 C=3 4 A=A^S";

    private readonly TcpListener _server = Started();
    private readonly TcpListener _bnls = Started();

    public BotEngineRapidReconnectTests() => LogonCheckCache.ClearForTests();

    public void Dispose()
    {
        _server.Stop();
        _bnls.Stop();
        LogonCheckCache.ClearForTests();
    }

    private static TcpListener Started()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        return listener;
    }

    private static int Port(TcpListener listener) => ((IPEndPoint)listener.LocalEndpoint).Port;

    private BotEngine Engine(string product = BncsProduct.Diablo, bool rapid = true, string? server = null, bool autoReconnect = false) => new(new BotConfig
    {
        Product = product,
        Username = "Tagban",
        Password = "huntertwo",
        BattlenetServer = server ?? "127.0.0.1",
        BattlenetPort = Port(_server),
        BnlsServer = "127.0.0.1",
        BnlsPort = Port(_bnls),
        RapidReconnect = rapid,
        AutoReconnect = autoReconnect,
        AutoReconnectDelaySeconds = 9999,
    });

    private static void CacheAnswers(string product = BncsProduct.Diablo)
    {
        LogonCheckCache.RememberVersionByte(product, 0x2A);
        LogonCheckCache.RememberVersionCheck(product, ChallengeTime, ChallengeFile, ChallengeFormula,
            new LogonCheckCache.VersionCheck(0x01000A0B, 0x12345678, "Diablo.exe 01/01/01 00:00:00 12345"));
    }

    private static List<string> Watch(BotEngine engine)
    {
        var logged = new List<string>();
        engine.Log += segments => { lock (logged) { logged.Add(string.Concat(segments.Select(s => s.Text))); } };
        return logged;
    }

    private static async Task<byte[]> ReadBncsPacketAsync(NetworkStream stream)
    {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(3));
        var header = new byte[4];
        await stream.ReadExactlyAsync(header, timeout.Token);
        var packet = new byte[header[2] | (header[3] << 8)];
        header.CopyTo(packet, 0);
        await stream.ReadExactlyAsync(packet.AsMemory(4), timeout.Token);
        return packet;
    }

    private static byte[] AuthInfoReply(string formula = ChallengeFormula) => new PacketWriter()
        .WriteDword(0) // logon type
        .WriteDword(0xCAFEF00D) // server token
        .WriteDword(0) // UDP value
        .WriteDword(ChallengeTime.Low)
        .WriteDword(ChallengeTime.High)
        .WriteNTString(ChallengeFile)
        .WriteNTString(formula)
        .ToBncsPacket(BncsPacketId.SID_AUTH_INFO);

    private static async Task<NetworkStream> AcceptBncsAsync(TcpListener server)
    {
        var client = await server.AcceptTcpClientAsync().WaitAsync(TimeSpan.FromSeconds(3));
        var stream = client.GetStream();
        var protocolByte = new byte[1];
        await stream.ReadExactlyAsync(protocolByte);
        Assert.Equal(0x01, protocolByte[0]);
        return stream;
    }

    [Fact(Timeout = 10000)]
    public async Task ARapidAttempt_WithCachedAnswers_LogsOnWithoutBnls()
    {
        CacheAnswers();
        await using var engine = Engine();
        var accepting = AcceptBncsAsync(_server);

        await (Task)typeof(BotEngine).GetMethod("ConnectForRapidReconnectAsync", Private)!.Invoke(engine, [CancellationToken.None])!;
        var stream = await accepting;

        var authInfo = await ReadBncsPacketAsync(stream);
        Assert.Equal((byte)BncsPacketId.SID_AUTH_INFO, authInfo[1]);
        Assert.Equal(0x2Au, new PacketReader(authInfo, offset: 16).ReadDword()); // the cached version byte

        await stream.WriteAsync(AuthInfoReply());
        var authCheck = await ReadBncsPacketAsync(stream);
        Assert.Equal((byte)BncsPacketId.SID_AUTH_CHECK, authCheck[1]);
        var reader = new PacketReader(authCheck, offset: 4);
        reader.ReadDword(); // client token
        Assert.Equal(0x01000A0Bu, reader.ReadDword());
        Assert.Equal(0x12345678u, reader.ReadDword());
        Assert.Equal(0u, reader.ReadDword()); // Diablo: no CD keys
        reader.ReadDword(); // spawn
        Assert.Equal("Diablo.exe 01/01/01 00:00:00 12345", reader.ReadNTString());

        Assert.False(_bnls.Pending(), "BNLS shouldn't have been asked anything");
    }

    [Fact(Timeout = 10000)]
    public async Task ARapidAttempt_WhoseServerChangedItsChallenge_LogsOnTheNormalWay()
    {
        CacheAnswers();
        await using var engine = Engine();
        var accepting = AcceptBncsAsync(_server);
        await (Task)typeof(BotEngine).GetMethod("ConnectForRapidReconnectAsync", Private)!.Invoke(engine, [CancellationToken.None])!;
        var stream = await accepting;
        await ReadBncsPacketAsync(stream);

        await stream.WriteAsync(AuthInfoReply(formula: "A=9 B=9 C=9 4 A=A+S"));

        using var bnlsClient = await _bnls.AcceptTcpClientAsync().WaitAsync(TimeSpan.FromSeconds(3));
        Assert.True(bnlsClient.Connected);
    }

    [Fact]
    public async Task BnlsAnswers_AreRememberedForTheNextLogon()
    {
        await using var engine = Engine();
        var auth = (AuthState)typeof(BotEngine).GetField("_auth", Private)!.GetValue(engine)!;
        auth.VersionCheckChallenge = (ChallengeTime, ChallengeFile, ChallengeFormula);
        var reply = new PacketWriter()
            .WriteDword(1) // success
            .WriteDword(0x01000A0B)
            .WriteDword(0x12345678)
            .WriteNTString("Diablo.exe")
            .WriteDword(0) // cookie
            .WriteDword(0) // version code
            .ToBnlsPacket(BnlsPacketId.BNLS_VERSIONCHECKEX2);

        await (Task)typeof(BotEngine).GetMethod("HandleVersionCheckEx2Async", Private)!.Invoke(engine, [reply])!;

        Assert.True(LogonCheckCache.TryGetVersionCheck(BncsProduct.Diablo, ChallengeTime, ChallengeFile, ChallengeFormula, out var check));
        Assert.Equal(new LogonCheckCache.VersionCheck(0x01000A0B, 0x12345678, "Diablo.exe"), check);
    }

    [Fact]
    public async Task CanLogOnWithoutBnls_OnlyForTheClassicLogonWithAnswersInHand()
    {
        await using var diablo = Engine();
        Assert.False(diablo.CanLogOnWithoutBnls()); // nothing cached yet
        CacheAnswers();
        Assert.True(diablo.CanLogOnWithoutBnls());

        CacheAnswers(BncsProduct.Warcraft3);
        await using var warcraft3 = Engine(BncsProduct.Warcraft3);
        Assert.False(warcraft3.CanLogOnWithoutBnls()); // its logon goes through BNLS regardless

        CacheAnswers(BncsProduct.Starcraft);
        await using var badKey = Engine(BncsProduct.Starcraft);
        badKey.Config.CdKey = "not a key";
        Assert.False(badKey.CanLogOnWithoutBnls());
    }

    [Fact(Timeout = 15000)]
    public async Task RapidReconnect_TriesEverySecondUntilStopped()
    {
        CacheAnswers();
        await using var engine = Engine();
        var accepted = new List<TimeSpan>();
        var clock = Stopwatch.StartNew();
        using var stop = new CancellationTokenSource();
        var server = Task.Run(async () =>
        {
            while (!stop.IsCancellationRequested)
            {
                try
                {
                    using var client = await _server.AcceptTcpClientAsync(stop.Token);
                    lock (accepted)
                    {
                        accepted.Add(clock.Elapsed);
                    }
                }
                catch (OperationCanceledException)
                {
                    return;
                }
            }
        });

        // The server hangs up on every attempt, like one that's still coming back up.
        typeof(BotEngine).GetMethod("MaybeScheduleAutoReconnect", Private)!.Invoke(engine, null);
        await Task.Delay(TimeSpan.FromMilliseconds(2400));
        await engine.DisconnectAsync();
        int countAtStop;
        lock (accepted)
        {
            countAtStop = accepted.Count;
        }

        await Task.Delay(TimeSpan.FromMilliseconds(1500));
        stop.Cancel();
        await server;

        Assert.InRange(countAtStop, 2, 4);
        Assert.InRange((accepted[1] - accepted[0]).TotalMilliseconds, 700, 1700);
        Assert.Equal(countAtStop, accepted.Count); // Disconnect stopped it
    }

    [Fact]
    public async Task RapidReconnect_IsNeverUsedOnOfficialBattlenet()
    {
        await using var engine = Engine(server: "useast.battle.net");
        var logged = Watch(engine);

        typeof(BotEngine).GetMethod("MaybeScheduleAutoReconnect", Private)!.Invoke(engine, null);
        await Task.Delay(100);

        Assert.False(engine.RapidReconnectApplies);
        Assert.Contains(logged, l => l.Contains("only for private servers", StringComparison.Ordinal));
        Assert.DoesNotContain(logged, l => l.Contains("trying every second", StringComparison.Ordinal));
    }

    [Fact]
    public async Task RapidReconnect_StillStandsDownAfterARejectedLogon()
    {
        await using var engine = Engine();
        var logged = Watch(engine);
        await (Task)typeof(BotEngine).GetMethod("HandleLogonResponse2Async", Private)!
            .Invoke(engine, [new PacketWriter().WriteDword(0x06).ToBncsPacket(BncsPacketId.SID_LOGONRESPONSE2)])!;

        typeof(BotEngine).GetMethod("MaybeScheduleAutoReconnect", Private)!.Invoke(engine, null);
        await Task.Delay(100);

        Assert.Contains(logged, l => l.Contains("Auto-reconnect skipped", StringComparison.Ordinal));
        Assert.DoesNotContain(logged, l => l.Contains("trying every second", StringComparison.Ordinal));
    }
}
