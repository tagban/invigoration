using System.Net;
using System.Net.Sockets;
using Invigoration.Sc2.Front;
using Invigoration.Sc2.Native;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Tests;

/// <summary>
/// Drives SunkenClient.ConnectAsync against a real loopback TCP listener
/// scripted to play the server side of Resume. This exercises everything a
/// live server test would — the actual socket connect, ResumeRequest
/// reaching the wire, and real Configuration/ProofRequest records being
/// received and parsed through RecordStream's buffering — except the final
/// thumbprint verification, which needs Blizzard's private key to produce a
/// genuine signature and so can't be exercised here (upstream's own test
/// suite has the identical gap for the same reason).
/// </summary>
public class SunkenClientIntegrationTests
{
    private static readonly byte[] Usage = "auth\0\0\0\0"u8.ToArray();

    private static readonly byte[] ThumbprintIdentity =
    [
        0xd7, 0xe6, 0x62, 0x40, 0x80, 0xc1, 0xab, 0xa6, 0x6d, 0xee, 0x63, 0xa6, 0xf3, 0x92, 0x8d, 0x8a,
        0x54, 0x69, 0x25, 0x7f, 0x58, 0x20, 0xb5, 0x72, 0x1f, 0xb8, 0xc3, 0x2b, 0x6b, 0x5b, 0xef, 0x5d,
    ];

    private static readonly byte[] SessionProofIdentity =
    [
        0x89, 0x50, 0x05, 0x34, 0x0a, 0x63, 0x0a, 0x64, 0x65, 0xa6, 0x5f, 0xec, 0x96, 0x32, 0x3c, 0x31,
        0x0b, 0xca, 0x8a, 0x9f, 0x66, 0xec, 0xee, 0xb1, 0x88, 0x7a, 0x9d, 0x6c, 0x0e, 0x67, 0x61, 0x2e,
    ];

    [Fact]
    public async Task ConnectAsync_ReachesThumbprintVerification_OverARealSocket()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;

        var serverClientTask = RunFakeServerAsync(listener);

        var handoff = new SunkenHandoff(
            Address: $"127.0.0.1:{port}",
            SessionKey: Enumerable.Range(0, 64).Select(i => (byte)i).ToArray(),
            AccountRegion: 1,
            GameAccountName: "Tagban",
            AccountMail: "player@example.com",
            LogonResponse: BuildLogonResponse3(accountRegion: 1, gameAccountRegion: 1, gameAccountName: "Tagban"));

        var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => SunkenClient.ConnectAsync(handoff, port));
        Assert.Contains("thumbprint proof failed", ex.Message);

        // Only dispose the server-side socket after the client has already consumed
        // everything it needed and thrown — disposing earlier risks a connection reset
        // racing the client's read of the last buffered bytes.
        (await serverClientTask).Dispose();
        listener.Stop();
    }

    private static byte[] BuildLogonResponse3(byte accountRegion, byte gameAccountRegion, string gameAccountName)
    {
        var writer = new BitWriter();
        writer.Write(0, 1); // Logon (0 bits) + m_result selector: success
        writer.Write(0, 3); // m_finalRequest: 0 modules
        writer.Write((uint)30, 32); // m_pingTimeout
        writer.Write(0, 1); // m_regulatorRules: absent
        writer.Write(0, 6); // m_givenName: 0 bytes
        writer.WriteBytes([], aligned: true);
        writer.Write(0, 6); // m_surname: 0 bytes
        writer.WriteBytes([], aligned: true);
        writer.Write(1, 32); // m_accountId
        writer.Write(accountRegion, 8);
        writer.Write(0UL, 64); // m_accountFlags
        writer.Write(gameAccountRegion, 8);
        var nameBytes = System.Text.Encoding.UTF8.GetBytes(gameAccountName);
        writer.Write((ulong)(nameBytes.Length - 1), 5);
        writer.WriteBytes(nameBytes, aligned: true);
        writer.Write(0UL, 64); // m_gameAccountFlags
        writer.Write(0, 32); // m_logonFailures
        writer.Align();
        return writer.ToBytes();
    }

    private static async Task<TcpClient> RunFakeServerAsync(TcpListener listener)
    {
        // This fake server never reads the client's ResumeRequest bytes — the OS buffers
        // them regardless, and only the client's outbound framing is under test here.
        var client = await listener.AcceptTcpClientAsync();
        var stream = client.GetStream();

        await stream.WriteAsync(EncodeConfiguration());
        await stream.WriteAsync(EncodeProofRequest());
        return client;
    }

    private static byte[] EncodeConfiguration()
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, commandId: 18, serviceSlot: ResumeHandshake.AuthenticationSlot);
        writer.Write(1, 1); // use_s3_depot
        writer.Align();
        return writer.ToBytes();
    }

    private static byte[] EncodeProofRequest()
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, commandId: 2, serviceSlot: ResumeHandshake.AuthenticationSlot);
        writer.Write(2, 3); // two modules

        // Thumbprint module: a 512-byte all-zero "signature" — structurally valid, cryptographically fake.
        writer.WriteBytes(Usage, aligned: true);
        writer.WriteBytes(ThumbprintIdentity, aligned: true);
        writer.Write(512, 10);
        writer.WriteBytes(new byte[512], aligned: true);

        // Session-proof module: phase 0 + a 16-byte server nonce.
        writer.WriteBytes(Usage, aligned: true);
        writer.WriteBytes(SessionProofIdentity, aligned: true);
        writer.Write(17, 10);
        writer.WriteBytes([0, .. new byte[16]], aligned: true);

        writer.Align();
        return writer.ToBytes();
    }
}
