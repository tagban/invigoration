using System.Net;
using System.Net.Sockets;
using Invigoration.Core;
using Invigoration.Core.Auth;
using Invigoration.Core.Config;
using Invigoration.Core.Protocol;
using System.Linq;

namespace Invigoration.Core.Tests;

/// <summary>
/// Drives BotEngine against a real loopback TCP listener standing in for a
/// BNLS server, to verify the framing/read-loop plumbing in FramedTcpClient
/// and the early handshake steps work over real sockets end-to-end, not just
/// in isolated unit tests of the packet math.
/// </summary>
public class BotEngineBnlsHandshakeTests
{
    [Fact(Timeout = 5000)]
    public async Task ConnectAsync_CompletesBnlsAuthorizeAndRequestsVersionByte()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;

        var config = new BotConfig
        {
            Username = "testuser",
            Password = "testpass",
            CdKey = "0000000000000000",
            BnlsServer = "127.0.0.1",
            BnlsPort = port,
            BattlenetServer = "127.0.0.1",
            BattlenetPort = 1, // nothing listens here; BNCS connect is expected to fail in this test
            Product = BncsProduct.DiabloII,
        };

        await using var engine = new BotEngine(config);
        var logs = new List<string>();
        engine.Log += segments => logs.Add(string.Concat(segments.Select(s => s.Text)));

        var connectTask = engine.ConnectAsync();
        using var server = await listener.AcceptTcpClientAsync();
        await using var serverStream = server.GetStream();

        // 1. Client sends BNLS_AUTHORIZE("Invigoration").
        var authorizePacket = await ReadBnlsPacketAsync(serverStream);
        Assert.Equal((byte)BnlsPacketId.BNLS_AUTHORIZE, authorizePacket[2]);
        var name = new PacketReader(authorizePacket, offset: 3).ReadNTString();
        Assert.Equal("Invigoration", name);

        // 2. Server sends a challenge; client must reply with the matching CRC32 checksum.
        const uint challenge = 0xABCD1234;
        var challengeReply = new PacketWriter().WriteDword(challenge).ToBnlsPacket(BnlsPacketId.BNLS_AUTHORIZE);
        await serverStream.WriteAsync(challengeReply);

        var proofPacket = await ReadBnlsPacketAsync(serverStream);
        Assert.Equal((byte)BnlsPacketId.BNLS_AUTHORIZEPROOF, proofPacket[2]);
        var response = new PacketReader(proofPacket, offset: 3).ReadDword();
        Assert.Equal(BnlsChecksum.Compute("Invigoration", challenge), response);

        // 3. Server accepts the proof; client requests the version byte.
        var proofReply = new PacketWriter().ToBnlsPacket(BnlsPacketId.BNLS_AUTHORIZEPROOF);
        await serverStream.WriteAsync(proofReply);

        var versionBytePacket = await ReadBnlsPacketAsync(serverStream);
        Assert.Equal((byte)BnlsPacketId.BNLS_REQUESTVERSIONBYTE, versionBytePacket[2]);
        var productByte = new PacketReader(versionBytePacket, offset: 3).ReadDword();
        Assert.Equal((uint)BncsProduct.GetBnlsProductByte(BncsProduct.DiabloII)!.Value, productByte);

        // 4. Server replies with a version byte; client then tries to connect to BNCS (which fails here by design).
        var versionByteReply = new PacketWriter()
            .WriteDword(1) // success
            .WriteDword(42) // version byte
            .ToBnlsPacket(BnlsPacketId.BNLS_REQUESTVERSIONBYTE);
        await serverStream.WriteAsync(versionByteReply);

        await Task.Delay(200); // let the engine process the reply and attempt the (failing) BNCS connect
        Assert.Contains(logs, m => m.Contains("Battle.net connecting to", StringComparison.Ordinal));

        await connectTask;
    }

    /// <summary>
    /// Regression test: BNLS_REQUESTVERSIONBYTE's reply payload is two
    /// DWORDs (the product byte echoed back, then the actual version byte),
    /// not a single DWORD. A prior bug read the first DWORD (the echo)
    /// instead of the second (the real value), which real BNCS servers
    /// silently rejected further down the handshake with no error — this
    /// pins the fix down at the source instead of relying on that symptom.
    /// </summary>
    [Fact(Timeout = 5000)]
    public async Task RequestVersionByteReply_SecondDwordIsUsedAsVersionByte_NotTheEchoedProductByte()
    {
        using var bnlsListener = new TcpListener(IPAddress.Loopback, 0);
        bnlsListener.Start();
        var bnlsPort = ((IPEndPoint)bnlsListener.LocalEndpoint).Port;

        using var bncsListener = new TcpListener(IPAddress.Loopback, 0);
        bncsListener.Start();
        var bncsPort = ((IPEndPoint)bncsListener.LocalEndpoint).Port;

        var config = new BotConfig
        {
            Username = "testuser",
            Password = "testpass",
            CdKey = "0000000000000000",
            BnlsServer = "127.0.0.1",
            BnlsPort = bnlsPort,
            BattlenetServer = "127.0.0.1",
            BattlenetPort = bncsPort,
            Product = BncsProduct.DiabloII,
        };

        await using var engine = new BotEngine(config);

        var connectTask = engine.ConnectAsync();
        using var bnlsServer = await bnlsListener.AcceptTcpClientAsync();
        await using var bnlsStream = bnlsServer.GetStream();

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_AUTHORIZE
        await bnlsStream.WriteAsync(new PacketWriter().WriteDword(0).ToBnlsPacket(BnlsPacketId.BNLS_AUTHORIZE));

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_AUTHORIZEPROOF
        await bnlsStream.WriteAsync(new PacketWriter().ToBnlsPacket(BnlsPacketId.BNLS_AUTHORIZEPROOF));

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_REQUESTVERSIONBYTE
        const uint echoedProductByte = 4;
        const uint realVersionByte = 0xAB;
        await bnlsStream.WriteAsync(
            new PacketWriter()
                .WriteDword(echoedProductByte)
                .WriteDword(realVersionByte)
                .ToBnlsPacket(BnlsPacketId.BNLS_REQUESTVERSIONBYTE));

        using var bncsServer = await bncsListener.AcceptTcpClientAsync();
        await using var bncsStream = bncsServer.GetStream();

        var protocolByte = new byte[1];
        await ReadExactAsync(bncsStream, protocolByte);
        Assert.Equal(0x01, protocolByte[0]);

        var authInfoPacket = await ReadBncsPacketAsync(bncsStream);
        var reader = new PacketReader(authInfoPacket, offset: 4);
        reader.Skip(4); // Protocol ID
        reader.Skip(8); // Platform ID + Product ID
        var versionByteSent = reader.ReadDword();

        Assert.Equal(realVersionByte, versionByteSent);

        await engine.DisconnectAsync();
        try
        {
            await connectTask;
        }
        catch
        {
            // Connection teardown races are irrelevant to this test's assertion.
        }
    }

    /// <summary>
    /// Covers a user-requested feature: overriding the version byte BNLS
    /// hands back, for private/PVPGN servers pinned to an older client
    /// version than BNLS's database assumes for a product. Same shape as the
    /// echoed-product-byte regression test above, but asserts the
    /// *configured* value reaches SID_AUTH_INFO, not BNLS's.
    /// </summary>
    [Theory(Timeout = 5000)]
    [InlineData("0x1A", 0x1Au)]
    [InlineData("26", 26u)]
    public async Task VersionByteOverride_WhenSet_IsUsedInsteadOfBnlsValue(string overrideText, uint expectedVersionByte)
    {
        using var bnlsListener = new TcpListener(IPAddress.Loopback, 0);
        bnlsListener.Start();
        var bnlsPort = ((IPEndPoint)bnlsListener.LocalEndpoint).Port;

        using var bncsListener = new TcpListener(IPAddress.Loopback, 0);
        bncsListener.Start();
        var bncsPort = ((IPEndPoint)bncsListener.LocalEndpoint).Port;

        var config = new BotConfig
        {
            Username = "testuser",
            Password = "testpass",
            CdKey = "0000000000000000",
            BnlsServer = "127.0.0.1",
            BnlsPort = bnlsPort,
            BattlenetServer = "127.0.0.1",
            BattlenetPort = bncsPort,
            Product = BncsProduct.DiabloII,
            VersionByteOverride = overrideText,
        };

        await using var engine = new BotEngine(config);

        var connectTask = engine.ConnectAsync();
        using var bnlsServer = await bnlsListener.AcceptTcpClientAsync();
        await using var bnlsStream = bnlsServer.GetStream();

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_AUTHORIZE
        await bnlsStream.WriteAsync(new PacketWriter().WriteDword(0).ToBnlsPacket(BnlsPacketId.BNLS_AUTHORIZE));

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_AUTHORIZEPROOF
        await bnlsStream.WriteAsync(new PacketWriter().ToBnlsPacket(BnlsPacketId.BNLS_AUTHORIZEPROOF));

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_REQUESTVERSIONBYTE
        const uint bnlsVersionByte = 0xFF; // deliberately different from expectedVersionByte, so the test fails if the override is ignored
        await bnlsStream.WriteAsync(
            new PacketWriter()
                .WriteDword(4) // echoed product byte
                .WriteDword(bnlsVersionByte)
                .ToBnlsPacket(BnlsPacketId.BNLS_REQUESTVERSIONBYTE));

        using var bncsServer = await bncsListener.AcceptTcpClientAsync();
        await using var bncsStream = bncsServer.GetStream();

        var protocolByte = new byte[1];
        await ReadExactAsync(bncsStream, protocolByte);

        var authInfoPacket = await ReadBncsPacketAsync(bncsStream);
        var reader = new PacketReader(authInfoPacket, offset: 4);
        reader.Skip(4); // Protocol ID
        reader.Skip(8); // Platform ID + Product ID
        var versionByteSent = reader.ReadDword();

        Assert.Equal(expectedVersionByte, versionByteSent);

        await engine.DisconnectAsync();
        try
        {
            await connectTask;
        }
        catch
        {
            // Connection teardown races are irrelevant to this test's assertion.
        }
    }

    /// <summary>
    /// Regression test for a real crash: BNLS_CDKEY_EX's per-key reply data
    /// repeats for each SUCCESSFULLY encrypted key (bnetdocs), not for each
    /// key requested. A live capture against a dual-CD-key (Warcraft III TFT)
    /// login had NumberRequested=2 but NumberSucceeded=1 (one of the two keys
    /// was rejected by BNLS) — the old handler looped NumberRequested times
    /// and threw IndexOutOfRangeException reading past the single 40-byte
    /// block actually present. The fix should recognize the shortfall and
    /// abort cleanly instead of crashing or sending an incomplete hash.
    /// </summary>
    [Fact(Timeout = 5000)]
    public async Task CdKeyExReply_FewerKeysSucceedThanRequested_AbortsWithoutCrashing()
    {
        using var bnlsListener = new TcpListener(IPAddress.Loopback, 0);
        bnlsListener.Start();
        var bnlsPort = ((IPEndPoint)bnlsListener.LocalEndpoint).Port;

        using var bncsListener = new TcpListener(IPAddress.Loopback, 0);
        bncsListener.Start();
        var bncsPort = ((IPEndPoint)bncsListener.LocalEndpoint).Port;

        var config = new BotConfig
        {
            Username = "testuser",
            Password = "testpass",
            CdKey = "12345678901234567890123456", // 26 chars: not 13/16, so CdKeyDecoder defers to BNLS
            ExpansionCdKey = "65432109876543210987654321",
            BnlsServer = "127.0.0.1",
            BnlsPort = bnlsPort,
            BattlenetServer = "127.0.0.1",
            BattlenetPort = bncsPort,
            Product = BncsProduct.Warcraft3TFT,
        };

        await using var engine = new BotEngine(config);
        var logs = new List<string>();
        engine.Log += segments => logs.Add(string.Concat(segments.Select(s => s.Text)));

        var connectTask = engine.ConnectAsync();
        using var bnlsServer = await bnlsListener.AcceptTcpClientAsync();
        await using var bnlsStream = bnlsServer.GetStream();

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_AUTHORIZE
        await bnlsStream.WriteAsync(new PacketWriter().WriteDword(0).ToBnlsPacket(BnlsPacketId.BNLS_AUTHORIZE));

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_AUTHORIZEPROOF
        await bnlsStream.WriteAsync(new PacketWriter().ToBnlsPacket(BnlsPacketId.BNLS_AUTHORIZEPROOF));

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_REQUESTVERSIONBYTE
        await bnlsStream.WriteAsync(
            new PacketWriter().WriteDword(0).WriteDword(42).ToBnlsPacket(BnlsPacketId.BNLS_REQUESTVERSIONBYTE));

        using var bncsServer = await bncsListener.AcceptTcpClientAsync();
        await using var bncsStream = bncsServer.GetStream();

        var protocolByte = new byte[1];
        await ReadExactAsync(bncsStream, protocolByte);

        await ReadBncsPacketAsync(bncsStream); // SID_AUTH_INFO request
        var authInfoReply = new PacketWriter()
            .WriteDword(0) // Logon Type
            .WriteDword(0xABCD) // Server token
            .WriteDword(0) // UDP value
            .WriteDword(0).WriteDword(0) // MPQ file time
            .WriteNTString("ver-IX86-1.mpq")
            .WriteNTString("A=A B=B C=C 4 A=A B=B C=C 4 A=A B=B C=C 4 A=A B=B C=C 4 A=A B=B C=C 4 A=A B=B C=C 4")
            .ToBncsPacket(BncsPacketId.SID_AUTH_INFO);
        await bncsStream.WriteAsync(authInfoReply);

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_VERSIONCHECKEX2 request
        var versionCheckReply = new PacketWriter()
            .WriteDword(1) // Success
            .WriteDword(1) // Exe version
            .WriteDword(2) // Exe checksum
            .WriteNTString("war3.exe 01/01/26 00:00:00 123456")
            .WriteDword(0) // Cookie, unused
            .WriteDword(0) // Version code, unused
            .ToBnlsPacket(BnlsPacketId.BNLS_VERSIONCHECKEX2);
        await bnlsStream.WriteAsync(versionCheckReply);

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_CDKEY_EX request

        // NumberRequested=2, NumberSucceeded=1, BitMask=0x1 (only the first
        // key succeeded), followed by exactly ONE 40-byte per-key block —
        // matching the real capture that crashed the old loop-by-requested code.
        var cdKeyExReply = new PacketWriter()
            .WriteDword(0) // Cookie
            .WriteByte(2) // Number requested
            .WriteByte(1) // Number succeeded
            .WriteDword(1) // Bit mask: only key 0 succeeded
            .WriteDword(0x1A2B3C4D) // Client session key for the one successful key
            .WriteBytes(new byte[36])
            .ToBnlsPacket(BnlsPacketId.BNLS_CDKEY_EX);
        await bnlsStream.WriteAsync(cdKeyExReply);

        await Task.Delay(200); // let the engine process the reply
        Assert.Contains(logs, m => m.Contains("BNLS rejected", StringComparison.Ordinal) && m.Contains("expansion CD key", StringComparison.Ordinal));

        await engine.DisconnectAsync();
        try
        {
            await connectTask;
        }
        catch
        {
            // Connection teardown races are irrelevant to this test's assertion.
        }
    }

    /// <summary>
    /// Companion happy-path case for the regression above: when both requested
    /// keys succeed, the two 40-byte per-key blocks should be read in full and
    /// concatenated into a 72-byte hash sent on to SID_AUTH_CHECK, using the
    /// FIRST key's session key as the client token (matching the single-key
    /// BNLS_CDKEY handler's convention).
    /// </summary>
    [Fact(Timeout = 5000)]
    public async Task CdKeyExReply_BothKeysSucceed_SendsCombinedHashToAuthCheck()
    {
        using var bnlsListener = new TcpListener(IPAddress.Loopback, 0);
        bnlsListener.Start();
        var bnlsPort = ((IPEndPoint)bnlsListener.LocalEndpoint).Port;

        using var bncsListener = new TcpListener(IPAddress.Loopback, 0);
        bncsListener.Start();
        var bncsPort = ((IPEndPoint)bncsListener.LocalEndpoint).Port;

        var config = new BotConfig
        {
            Username = "testuser",
            Password = "testpass",
            CdKey = "12345678901234567890123456",
            ExpansionCdKey = "65432109876543210987654321",
            BnlsServer = "127.0.0.1",
            BnlsPort = bnlsPort,
            BattlenetServer = "127.0.0.1",
            BattlenetPort = bncsPort,
            Product = BncsProduct.Warcraft3TFT,
        };

        await using var engine = new BotEngine(config);

        var connectTask = engine.ConnectAsync();
        using var bnlsServer = await bnlsListener.AcceptTcpClientAsync();
        await using var bnlsStream = bnlsServer.GetStream();

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_AUTHORIZE
        await bnlsStream.WriteAsync(new PacketWriter().WriteDword(0).ToBnlsPacket(BnlsPacketId.BNLS_AUTHORIZE));

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_AUTHORIZEPROOF
        await bnlsStream.WriteAsync(new PacketWriter().ToBnlsPacket(BnlsPacketId.BNLS_AUTHORIZEPROOF));

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_REQUESTVERSIONBYTE
        await bnlsStream.WriteAsync(
            new PacketWriter().WriteDword(0).WriteDword(42).ToBnlsPacket(BnlsPacketId.BNLS_REQUESTVERSIONBYTE));

        using var bncsServer = await bncsListener.AcceptTcpClientAsync();
        await using var bncsStream = bncsServer.GetStream();

        var protocolByte = new byte[1];
        await ReadExactAsync(bncsStream, protocolByte);

        await ReadBncsPacketAsync(bncsStream); // SID_AUTH_INFO request
        var authInfoReply = new PacketWriter()
            .WriteDword(0)
            .WriteDword(0xABCD)
            .WriteDword(0)
            .WriteDword(0).WriteDword(0)
            .WriteNTString("ver-IX86-1.mpq")
            .WriteNTString("A=A B=B C=C 4 A=A B=B C=C 4 A=A B=B C=C 4 A=A B=B C=C 4 A=A B=B C=C 4 A=A B=B C=C 4")
            .ToBncsPacket(BncsPacketId.SID_AUTH_INFO);
        await bncsStream.WriteAsync(authInfoReply);

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_VERSIONCHECKEX2 request
        var versionCheckReply = new PacketWriter()
            .WriteDword(1)
            .WriteDword(1)
            .WriteDword(2)
            .WriteNTString("war3.exe 01/01/26 00:00:00 123456")
            .WriteDword(0)
            .WriteDword(0)
            .ToBnlsPacket(BnlsPacketId.BNLS_VERSIONCHECKEX2);
        await bnlsStream.WriteAsync(versionCheckReply);

        await ReadBnlsPacketAsync(bnlsStream); // BNLS_CDKEY_EX request

        var firstKeyHash = Enumerable.Repeat((byte)0xAA, 36).ToArray();
        var secondKeyHash = Enumerable.Repeat((byte)0xBB, 36).ToArray();
        var cdKeyExReply = new PacketWriter()
            .WriteDword(0)
            .WriteByte(2)
            .WriteByte(2) // both keys succeeded
            .WriteDword(3) // bit mask: both bits set
            .WriteDword(0x11111111).WriteBytes(firstKeyHash)
            .WriteDword(0x22222222).WriteBytes(secondKeyHash)
            .ToBnlsPacket(BnlsPacketId.BNLS_CDKEY_EX);
        await bnlsStream.WriteAsync(cdKeyExReply);

        var authCheckPacket = await ReadBncsPacketAsync(bncsStream);
        var reader = new PacketReader(authCheckPacket, offset: 4);
        var clientToken = reader.ReadDword();
        reader.Skip(4); // Exe version
        reader.Skip(4); // Exe checksum
        var numKeys = reader.ReadDword();
        reader.Skip(4); // spawn
        var combinedHash = reader.ReadRaw(72);

        Assert.Equal(0x11111111u, clientToken);
        Assert.Equal(2u, numKeys);
        Assert.Equal(firstKeyHash.Concat(secondKeyHash), combinedHash);

        await engine.DisconnectAsync();
        try
        {
            await connectTask;
        }
        catch
        {
            // Connection teardown races are irrelevant to this test's assertion.
        }
    }

    private static async Task<byte[]> ReadBncsPacketAsync(NetworkStream stream)
    {
        var header = new byte[4];
        await ReadExactAsync(stream, header).ConfigureAwait(false);
        var length = header[2] | (header[3] << 8);
        var packet = new byte[length];
        Array.Copy(header, packet, 4);
        await ReadExactAsync(stream, packet.AsMemory(4)).ConfigureAwait(false);
        return packet;
    }

    private static async Task<byte[]> ReadBnlsPacketAsync(NetworkStream stream)
    {
        var header = new byte[3];
        await ReadExactAsync(stream, header).ConfigureAwait(false);
        var length = header[0] | (header[1] << 8);
        var packet = new byte[length];
        Array.Copy(header, packet, 3);
        await ReadExactAsync(stream, packet.AsMemory(3)).ConfigureAwait(false);
        return packet;
    }

    private static async Task ReadExactAsync(NetworkStream stream, Memory<byte> buffer)
    {
        var totalRead = 0;
        while (totalRead < buffer.Length)
        {
            var read = await stream.ReadAsync(buffer[totalRead..]).ConfigureAwait(false);
            if (read == 0)
            {
                throw new IOException("Stream closed before expected bytes were read.");
            }

            totalRead += read;
        }
    }
}
