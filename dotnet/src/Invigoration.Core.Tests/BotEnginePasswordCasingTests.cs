using System.Net;
using System.Net.Sockets;
using System.Reflection;
using Invigoration.Core.Auth;
using Invigoration.Core.Config;
using Invigoration.Core.Protocol;

namespace Invigoration.Core.Tests;

/// <summary>
/// The logon's password casing: sent as typed first, retried once in the game clients' casing
/// when the server calls it incorrect, and the "Normalize Password" change that makes the
/// account itself lowercase. Handlers are driven directly (same reflection approach as
/// BotEngineJoinBurstTests) against a loopback BNLS listener, so the retry's actual outgoing
/// BNLS request can be read back and checked byte for byte.
/// </summary>
public class BotEnginePasswordCasingTests
{
    private sealed class Harness : IAsyncDisposable
    {
        private readonly TcpListener _listener;
        private TcpClient? _server;

        public BotEngine Engine { get; }
        public BotConfig Config { get; }
        public List<string> Logs { get; } = [];
        public NetworkStream Bnls { get; private set; } = null!;

        private Harness(TcpListener listener, BotConfig config)
        {
            _listener = listener;
            Config = config;
            Engine = new BotEngine(config);
            Engine.Log += segments => { lock (Logs) { Logs.Add(string.Concat(segments.Select(s => s.Text))); } };
        }

        public static async Task<Harness> StartAsync(string product, string password)
        {
            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            var config = new BotConfig
            {
                Username = "Tagban",
                Password = password,
                CdKey = "0000000000000000",
                BnlsServer = "127.0.0.1",
                BnlsPort = ((IPEndPoint)listener.LocalEndpoint).Port,
                BattlenetServer = "127.0.0.1",
                BattlenetPort = 1,
                Product = product,
            };

            var harness = new Harness(listener, config);
            await harness.AcceptBnlsAsync();
            return harness;
        }

        /// <summary>Connects the engine's BNLS side and walks the authorize exchange, stopping before the version-byte reply so nothing moves on to BNCS.</summary>
        public async Task AcceptBnlsAsync()
        {
            var connect = Engine.ConnectAsync();
            _server?.Dispose();
            _server = await _listener.AcceptTcpClientAsync();
            Bnls = _server.GetStream();
            await connect;

            await ReadBnlsPacketAsync(); // BNLS_AUTHORIZE
            await Bnls.WriteAsync(new PacketWriter().WriteDword(0).ToBnlsPacket(BnlsPacketId.BNLS_AUTHORIZE));
            await ReadBnlsPacketAsync(); // BNLS_AUTHORIZEPROOF
            await Bnls.WriteAsync(new PacketWriter().ToBnlsPacket(BnlsPacketId.BNLS_AUTHORIZEPROOF));
            await ReadBnlsPacketAsync(); // BNLS_REQUESTVERSIONBYTE — deliberately left unanswered
        }

        public AuthState Auth => (AuthState)typeof(BotEngine).GetField("_auth", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(Engine)!;

        public Task InvokeAsync(string handler, byte[] frame) =>
            (Task)typeof(BotEngine).GetMethod(handler, BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(Engine, [frame])!;

        public async Task<byte[]> ReadBnlsPacketAsync()
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(3));
            var header = new byte[3];
            await Bnls.ReadExactlyAsync(header, timeout.Token);
            var packet = new byte[header[0] | (header[1] << 8)];
            header.CopyTo(packet, 0);
            await Bnls.ReadExactlyAsync(packet.AsMemory(3), timeout.Token);
            return packet;
        }

        public bool Logged(string fragment)
        {
            lock (Logs)
            {
                return Logs.Any(l => l.Contains(fragment, StringComparison.Ordinal));
            }
        }

        public async ValueTask DisposeAsync()
        {
            await Engine.DisposeAsync();
            _server?.Dispose();
            _listener.Stop();
        }
    }

    private static byte[] LogonResponse2(uint status) => new PacketWriter().WriteDword(status).ToBncsPacket(BncsPacketId.SID_LOGONRESPONSE2);

    private static byte[] AccountLogonProof(uint status) => new PacketWriter().WriteDword(status).ToBncsPacket(BncsPacketId.SID_AUTH_ACCOUNTLOGONPROOF);

    private static byte[] ChangePasswordReply(bool success) => new PacketWriter().WriteDword(success ? 1u : 0u).ToBncsPacket(BncsPacketId.SID_CHANGEPASSWORD);

    [Fact(Timeout = 10000)]
    public async Task ClassicLogon_RejectedAsTyped_RetriesOnceInLowercase()
    {
        await using var h = await Harness.StartAsync(BncsProduct.Starcraft, "HunterTWO");
        h.Auth.LogonPassword = h.Config.Password;

        await h.InvokeAsync("HandleLogonResponse2Async", LogonResponse2(0x02));

        var retry = await h.ReadBnlsPacketAsync();
        Assert.Equal((byte)BnlsPacketId.BNLS_HASHDATA, retry[2]);
        var reader = new PacketReader(retry, offset: 3);
        Assert.Equal(9u, reader.ReadDword());
        Assert.Equal(0u, reader.ReadDword());
        Assert.Equal("huntertwo", System.Text.Encoding.ASCII.GetString(reader.ReadRaw(9)));
        Assert.Equal("huntertwo", h.Auth.LogonPassword);
        Assert.True(h.Logged("retrying in lowercase"));
        Assert.False(h.Logged("incorrect password"));

        // A second rejection is a genuinely wrong password: report it, don't loop.
        await h.InvokeAsync("HandleLogonResponse2Async", LogonResponse2(0x02));
        Assert.True(h.Logged("incorrect password"));
    }

    [Fact(Timeout = 10000)]
    public async Task NlsLogon_RejectedAsTyped_RetriesOnceInUppercase()
    {
        await using var h = await Harness.StartAsync(BncsProduct.Warcraft3TFT, "HunterTWO");
        h.Auth.LogonPassword = h.Config.Password;

        await h.InvokeAsync("HandleAuthAccountLogonProofAsync", AccountLogonProof(0x02));

        var retry = await h.ReadBnlsPacketAsync();
        Assert.Equal((byte)BnlsPacketId.BNLS_LOGONCHALLENGE, retry[2]);
        var reader = new PacketReader(retry, offset: 3);
        Assert.Equal("Tagban", reader.ReadNTString());
        Assert.Equal("HUNTERTWO", reader.ReadNTString());
        Assert.True(h.Logged("retrying in uppercase"));
    }

    [Fact(Timeout = 10000)]
    public async Task ClassicLogon_AlreadyLowercase_ReportsWithoutRetrying()
    {
        await using var h = await Harness.StartAsync(BncsProduct.DiabloII, "huntertwo");
        h.Auth.LogonPassword = h.Config.Password;

        await h.InvokeAsync("HandleLogonResponse2Async", LogonResponse2(0x02));

        Assert.True(h.Logged("incorrect password"));
        Assert.False(h.Auth.RetriedPasswordCasing);
    }

    [Theory(Timeout = 10000)]
    [InlineData(BncsProduct.Warcraft3TFT, "HunterTWO", "different way")]
    [InlineData(BncsProduct.Starcraft, "huntertwo", "already all lowercase")]
    [InlineData(BncsProduct.Starcraft, "", "Set this bot's username and password")]
    public async Task NormalizePassword_RefusesWhenItCantApply(string product, string password, string expectedReason)
    {
        await using var engine = new BotEngine(new BotConfig { Username = "Tagban", Password = password, Product = product });
        var logs = new List<string>();
        engine.Log += segments => logs.Add(string.Concat(segments.Select(s => s.Text)));

        Assert.Contains(expectedReason, engine.NormalizePasswordUnavailableReason);
        await engine.NormalizePasswordAsync();

        Assert.Contains(logs, l => l.Contains(expectedReason, StringComparison.Ordinal));
        Assert.DoesNotContain(logs, l => l.Contains("reconnecting to change it", StringComparison.Ordinal));
    }

    [Fact(Timeout = 10000)]
    public async Task ChangePasswordReply_Success_SavesTheLowercasePasswordAndReconnects()
    {
        await using var h = await Harness.StartAsync(BncsProduct.Starcraft, "HunterTWO");
        var persisted = 0;
        h.Engine.ConfigPersistNeeded += () => persisted++;
        h.Auth.NewPassword = "huntertwo";

        var reply = h.InvokeAsync("HandleChangePasswordReplyAsync", ChangePasswordReply(true));
        await h.AcceptBnlsAsync(); // the reconnect
        await reply;

        Assert.Equal("huntertwo", h.Config.Password);
        Assert.Equal(1, persisted);
        Assert.True(h.Logged("Password normalized"));
    }

    [Fact(Timeout = 10000)]
    public async Task ChangePasswordReply_Refused_LeavesThePasswordAlone()
    {
        await using var h = await Harness.StartAsync(BncsProduct.Starcraft, "HunterTWO");
        var persisted = 0;
        h.Engine.ConfigPersistNeeded += () => persisted++;
        h.Auth.NewPassword = "huntertwo";

        await h.InvokeAsync("HandleChangePasswordReplyAsync", ChangePasswordReply(false));

        Assert.Equal("HunterTWO", h.Config.Password);
        Assert.Equal(0, persisted);
        Assert.True(h.Logged("refused the password change"));
    }

    // The change must never stay armed for a later, unrelated Connect.
    [Fact(Timeout = 10000)]
    public async Task Disconnect_DisarmsAPendingPasswordChange()
    {
        await using var engine = new BotEngine(new BotConfig { Product = BncsProduct.Starcraft });
        var auth = (AuthState)typeof(BotEngine).GetField("_auth", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(engine)!;
        auth.ChangePasswordRequested = true;
        auth.NewPassword = "huntertwo";

        await engine.DisconnectAsync();

        Assert.False(auth.ChangePasswordRequested);
        Assert.Equal("", auth.NewPassword);
    }
}
