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
