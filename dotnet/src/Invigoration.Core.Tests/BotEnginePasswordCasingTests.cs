using System.Net;
using System.Net.Sockets;
using System.Reflection;
using Invigoration.Core.Auth;
using Invigoration.Core.Config;
using Invigoration.Core.Protocol;

namespace Invigoration.Core.Tests;

/// <summary>
/// The logon's password casing: sent in the game clients' casing first, tried once more exactly as
/// typed (on a fresh connection) when the server calls it incorrect, whichever worked remembered for
/// next time, and the "Normalize Password" change that makes the account itself lowercase. Handlers are driven directly (same reflection approach as
/// BotEngineJoinBurstTests) against loopback BNLS and BNCS listeners, so what the engine actually
/// sends can be read back and checked byte for byte — including that the password itself never
/// goes to BNLS, only its hashes to Battle.net.
/// </summary>
public class BotEnginePasswordCasingTests
{
    private sealed class Harness : IAsyncDisposable
    {
        private readonly TcpListener _listener;
        private readonly TcpListener _bncsListener = StartedListener();
        private TcpClient? _server;
        private TcpClient? _bncsServer;

        public BotEngine Engine { get; }
        public BotConfig Config { get; }
        public List<string> Logs { get; } = [];
        public NetworkStream Bnls { get; private set; } = null!;
        public NetworkStream Bncs { get; private set; } = null!;

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

        /// <summary>Accepts the engine's BNLS connection — starting it, unless the engine is already reconnecting on its own — and walks the authorize exchange, stopping before the version-byte reply so nothing moves on to BNCS.</summary>
        public async Task AcceptBnlsAsync(bool startConnect = true)
        {
            var connect = startConnect ? Engine.ConnectAsync() : Task.CompletedTask;
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

        /// <summary>Connects the engine's BNCS socket to a loopback listener and reads past what every connection opens with (protocol byte, SID_AUTH_INFO), so the next packet read is whatever the test provokes.</summary>
        public async Task ConnectBncsAsync()
        {
            var bncs = (Networking.FramedTcpClient)typeof(BotEngine).GetField("_bncs", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(Engine)!;
            var accept = _bncsListener.AcceptTcpClientAsync();
            await bncs.ConnectAsync("127.0.0.1", ((IPEndPoint)_bncsListener.LocalEndpoint).Port);
            _bncsServer = await accept;
            Bncs = _bncsServer.GetStream();

            var protocolByte = new byte[1];
            await Bncs.ReadExactlyAsync(protocolByte);
            Assert.Equal(0x01, protocolByte[0]);
            Assert.Equal((byte)BncsPacketId.SID_AUTH_INFO, (await ReadBncsPacketAsync())[1]);
        }

        public async Task<byte[]> ReadBncsPacketAsync()
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(3));
            var header = new byte[4];
            await Bncs.ReadExactlyAsync(header, timeout.Token);
            var packet = new byte[header[2] | (header[3] << 8)];
            header.CopyTo(packet, 0);
            await Bncs.ReadExactlyAsync(packet.AsMemory(4), timeout.Token);
            return packet;
        }

        /// <summary>Whether anything at all has been sent to BNLS since the authorize exchange.</summary>
        public async Task<bool> BnlsReceivedAnythingAsync()
        {
            await Task.Delay(150);
            return Bnls.DataAvailable;
        }

        private static TcpListener StartedListener()
        {
            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            return listener;
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
            _bncsServer?.Dispose();
            _listener.Stop();
            _bncsListener.Stop();
        }
    }

    private static byte[] LogonResponse2(uint status) => new PacketWriter().WriteDword(status).ToBncsPacket(BncsPacketId.SID_LOGONRESPONSE2);

    private static byte[] AccountLogonProof(uint status) => new PacketWriter().WriteDword(status).ToBncsPacket(BncsPacketId.SID_AUTH_ACCOUNTLOGONPROOF);

    private static byte[] ChangePasswordReply(bool success) => new PacketWriter().WriteDword(success ? 1u : 0u).ToBncsPacket(BncsPacketId.SID_CHANGEPASSWORD);

    private const uint ClientToken = 0x12345678;
    private const uint ServerToken = 0x9ABCDEF0;

    private static byte[] AuthCheckPassed() => new PacketWriter().WriteDword(0).WriteDword(0).ToBncsPacket(BncsPacketId.SID_AUTH_CHECK);

    /// <summary>Reads SID_UDPPINGRESPONSE and SID_GETICONDATA, then returns the logon packet's password proof (checking tokens and name on the way).</summary>
    private static async Task<byte[]> ReadLogonProofAsync(Harness h)
    {
        Assert.Equal((byte)BncsPacketId.SID_UDPPINGRESPONSE, (await h.ReadBncsPacketAsync())[1]);
        Assert.Equal((byte)BncsPacketId.SID_GETICONDATA, (await h.ReadBncsPacketAsync())[1]);
        var logon = await h.ReadBncsPacketAsync();
        Assert.Equal((byte)BncsPacketId.SID_LOGONRESPONSE2, logon[1]);
        var reader = new PacketReader(logon, offset: 4);
        Assert.Equal(ClientToken, reader.ReadDword());
        Assert.Equal(ServerToken, reader.ReadDword());
        var proof = reader.ReadRaw(20);
        Assert.Equal("Tagban", reader.ReadNTString());
        return proof;
    }

    private static async Task PassAuthCheckAsync(Harness h)
    {
        await h.ConnectBncsAsync();
        h.Auth.ClientToken = ClientToken;
        h.Auth.ServerToken = ServerToken;
        await h.InvokeAsync("HandleAuthCheckAsync", AuthCheckPassed());
    }

    [Fact(Timeout = 10000)]
    public async Task ClassicLogon_SendsTheGameClientsCasingFirst_HashedLocally_NothingToBnls()
    {
        await using var h = await Harness.StartAsync(BncsProduct.Warcraft2BNE, "HunterTWO");

        await PassAuthCheckAsync(h);

        Assert.Equal(BattlenetPassword.Proof(ClientToken, ServerToken, "huntertwo"), await ReadLogonProofAsync(h));
        Assert.False(await h.BnlsReceivedAnythingAsync());
    }

    [Fact(Timeout = 10000)]
    public async Task ClassicLogon_RememberedAsTyped_SendsItAsTypedFirst()
    {
        await using var h = await Harness.StartAsync(BncsProduct.Starcraft, "HunterTWO");
        h.Config.PasswordSentAsTyped = true;

        await PassAuthCheckAsync(h);

        Assert.Equal(BattlenetPassword.Proof(ClientToken, ServerToken, "HunterTWO"), await ReadLogonProofAsync(h));
    }

    [Fact(Timeout = 10000)]
    public async Task ClassicLogon_Rejected_TriesAsTypedOnAFreshConnection_AndRemembersWhatWorked()
    {
        await using var h = await Harness.StartAsync(BncsProduct.Starcraft, "HunterTWO");
        var persisted = 0;
        h.Engine.ConfigPersistNeeded += () => persisted++;
        await PassAuthCheckAsync(h);
        await ReadLogonProofAsync(h); // lowercase

        // Rejected: the engine disconnects and starts a whole new connection (BNLS first, as always).
        var rejected = h.InvokeAsync("HandleLogonResponse2Async", LogonResponse2(0x02));
        await h.AcceptBnlsAsync(startConnect: false);
        await rejected;
        Assert.True(h.Logged("reconnecting to try it exactly as typed"));
        Assert.False(h.Logged("incorrect password"));

        // That connection's logon is the password as typed.
        await PassAuthCheckAsync(h);
        Assert.Equal(BattlenetPassword.Proof(ClientToken, ServerToken, "HunterTWO"), await ReadLogonProofAsync(h));
        Assert.Equal("HunterTWO", h.Auth.LogonPassword);

        // It gets in: remembered, so the next logon starts as typed.
        await h.InvokeAsync("HandleLogonResponse2Async", LogonResponse2(0x00));
        Assert.True(h.Config.PasswordSentAsTyped);
        Assert.Equal(1, persisted);
        Assert.True(h.Logged("Remembered"));
    }

    [Fact(Timeout = 10000)]
    public async Task ClassicLogon_RejectedBothWays_ReportsTheWrongPasswordWithoutLooping()
    {
        await using var h = await Harness.StartAsync(BncsProduct.Starcraft, "HunterTWO");
        h.Auth.RetriedPasswordCasing = true; // this connection already is the retry

        await h.InvokeAsync("HandleLogonResponse2Async", LogonResponse2(0x02));

        Assert.True(h.Logged("incorrect password"));
        Assert.False(h.Config.PasswordSentAsTyped);
    }

    [Fact(Timeout = 10000)]
    public async Task ClassicLogon_FirstTryGetsIn_ChangesNothing()
    {
        await using var h = await Harness.StartAsync(BncsProduct.Starcraft, "HunterTWO");
        var persisted = 0;
        h.Engine.ConfigPersistNeeded += () => persisted++;

        await h.InvokeAsync("HandleLogonResponse2Async", LogonResponse2(0x00));

        Assert.False(h.Config.PasswordSentAsTyped);
        Assert.Equal(0, persisted);
    }

    [Fact(Timeout = 10000)]
    public async Task NormalizePassword_SendsTheOldProofAndTheNewHash_HashedLocally()
    {
        await using var h = await Harness.StartAsync(BncsProduct.Starcraft, "HunterTWO");
        await h.ConnectBncsAsync();
        h.Auth.ClientToken = ClientToken;
        h.Auth.ServerToken = ServerToken;
        h.Auth.ChangePasswordRequested = true;
        h.Auth.NewPassword = "huntertwo";

        await h.InvokeAsync("HandleAuthCheckAsync", AuthCheckPassed());

        await h.ReadBncsPacketAsync(); // SID_UDPPINGRESPONSE
        await h.ReadBncsPacketAsync(); // SID_GETICONDATA
        var change = await h.ReadBncsPacketAsync();
        Assert.Equal((byte)BncsPacketId.SID_CHANGEPASSWORD, change[1]);
        var reader = new PacketReader(change, offset: 4);
        Assert.Equal(ClientToken, reader.ReadDword());
        Assert.Equal(ServerToken, reader.ReadDword());
        Assert.Equal(BattlenetPassword.Proof(ClientToken, ServerToken, "HunterTWO"), reader.ReadRaw(20));
        Assert.Equal(BattlenetPassword.Hash("huntertwo"), reader.ReadRaw(20));
        Assert.Equal("Tagban", reader.ReadNTString());
        Assert.False(h.Auth.ChangePasswordRequested);
        Assert.False(await h.BnlsReceivedAnythingAsync());
    }

    [Fact(Timeout = 10000)]
    public async Task NlsLogon_SendsUppercaseFirst_AndARejectionReconnectsToTryAsTyped()
    {
        await using var h = await Harness.StartAsync(BncsProduct.Warcraft3TFT, "HunterTWO");

        await h.InvokeAsync("HandleAuthCheckAsync", AuthCheckPassed());
        var challenge = await h.ReadBnlsPacketAsync();
        Assert.Equal((byte)BnlsPacketId.BNLS_LOGONCHALLENGE, challenge[2]);
        var reader = new PacketReader(challenge, offset: 3);
        Assert.Equal("Tagban", reader.ReadNTString());
        Assert.Equal("HUNTERTWO", reader.ReadNTString());

        var rejected = h.InvokeAsync("HandleAuthAccountLogonProofAsync", AccountLogonProof(0x02));
        await h.AcceptBnlsAsync(startConnect: false);
        await rejected;
        Assert.True(h.Logged("reconnecting to try it exactly as typed"));

        await h.InvokeAsync("HandleAuthCheckAsync", AuthCheckPassed());
        var retry = await h.ReadBnlsPacketAsync();
        reader = new PacketReader(retry, offset: 3);
        Assert.Equal("Tagban", reader.ReadNTString());
        Assert.Equal("HunterTWO", reader.ReadNTString());
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
        h.Config.PasswordSentAsTyped = true;
        var persisted = 0;
        h.Engine.ConfigPersistNeeded += () => persisted++;
        h.Auth.NewPassword = "huntertwo";

        var reply = h.InvokeAsync("HandleChangePasswordReplyAsync", ChangePasswordReply(true));
        await h.AcceptBnlsAsync(startConnect: false); // the engine's own reconnect
        await reply;

        Assert.Equal("huntertwo", h.Config.Password);
        Assert.False(h.Config.PasswordSentAsTyped);
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
