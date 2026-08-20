using System.Net;
using System.Net.Sockets;
using Invigoration.Core.Networking;

namespace Invigoration.Core.Tests;

/// <summary>
/// Exercises the SOCKS5/HTTP CONNECT handshakes against a real local
/// TcpListener standing in for a proxy — verifies the actual bytes on the
/// wire round-trip correctly, not just that the pure builder/parser
/// functions individually look right.
/// </summary>
public class ProxyConnectorTests
{
    private static (TcpListener Listener, int Port) StartFakeProxy()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        return (listener, ((IPEndPoint)listener.LocalEndpoint).Port);
    }

    [Fact]
    public async Task Socks5_NoAuth_SuccessfulConnect_CompletesNegotiation()
    {
        var (listener, port) = StartFakeProxy();
        using var _ = listener;

        var serverTask = Task.Run(async () =>
        {
            using var proxyClient = await listener.AcceptTcpClientAsync();
            using var proxyStream = proxyClient.GetStream();

            var greeting = await ReadExactAsync(proxyStream, 3); // VER, NMETHODS=1, METHOD=0x00
            Assert.Equal([0x05, 0x01, 0x00], greeting);
            await proxyStream.WriteAsync(new byte[] { 0x05, 0x00 }); // no-auth selected

            var connectHeader = await ReadExactAsync(proxyStream, 5); // VER, CMD, RSV, ATYP, domain-length
            Assert.Equal(0x05, connectHeader[0]);
            Assert.Equal(0x01, connectHeader[1]); // CONNECT
            Assert.Equal(0x03, connectHeader[3]); // domain name
            var domainLength = connectHeader[4];
            await ReadExactAsync(proxyStream, domainLength + 2); // domain + port

            // Reply: success, bound address 0.0.0.0:0 (IPv4)
            await proxyStream.WriteAsync(new byte[] { 0x05, 0x00, 0x00, 0x01, 0, 0, 0, 0, 0, 0 });
        });

        using var client = new TcpClient();
        await client.ConnectAsync(IPAddress.Loopback, port);
        using var stream = client.GetStream();

        await Socks5Connector.NegotiateAsync(stream, "target.example.com", 6112, null, null, CancellationToken.None);
        await serverTask;
    }

    [Fact]
    public async Task Socks5_UsernamePasswordAuth_SuccessfulConnect_CompletesNegotiation()
    {
        var (listener, port) = StartFakeProxy();
        using var _ = listener;

        var serverTask = Task.Run(async () =>
        {
            using var proxyClient = await listener.AcceptTcpClientAsync();
            using var proxyStream = proxyClient.GetStream();

            var greeting = await ReadExactAsync(proxyStream, 4); // VER, NMETHODS=2, 0x00, 0x02
            Assert.Equal(0x05, greeting[0]);
            await proxyStream.WriteAsync(new byte[] { 0x05, 0x02 }); // username/password selected

            var authHeader = await ReadExactAsync(proxyStream, 2); // VER=1, ULEN
            var uname = await ReadExactAsync(proxyStream, authHeader[1]);
            Assert.Equal("botuser", System.Text.Encoding.UTF8.GetString(uname));
            var plenByte = await ReadExactAsync(proxyStream, 1);
            var pass = await ReadExactAsync(proxyStream, plenByte[0]);
            Assert.Equal("botpass", System.Text.Encoding.UTF8.GetString(pass));
            await proxyStream.WriteAsync(new byte[] { 0x01, 0x00 }); // auth success

            var connectHeader = await ReadExactAsync(proxyStream, 5);
            var domainLength = connectHeader[4];
            await ReadExactAsync(proxyStream, domainLength + 2);
            await proxyStream.WriteAsync(new byte[] { 0x05, 0x00, 0x00, 0x01, 0, 0, 0, 0, 0, 0 });
        });

        using var client = new TcpClient();
        await client.ConnectAsync(IPAddress.Loopback, port);
        using var stream = client.GetStream();

        await Socks5Connector.NegotiateAsync(stream, "target.example.com", 6112, "botuser", "botpass", CancellationToken.None);
        await serverTask;
    }

    [Fact]
    public async Task Socks5_ProxyRefusesConnection_ThrowsWithReplyCodeDescription()
    {
        var (listener, port) = StartFakeProxy();
        using var _ = listener;

        var serverTask = Task.Run(async () =>
        {
            using var proxyClient = await listener.AcceptTcpClientAsync();
            using var proxyStream = proxyClient.GetStream();
            await ReadExactAsync(proxyStream, 3);
            await proxyStream.WriteAsync(new byte[] { 0x05, 0x00 });
            var connectHeader = await ReadExactAsync(proxyStream, 5);
            await ReadExactAsync(proxyStream, connectHeader[4] + 2);
            // Reply: connection refused (0x05)
            await proxyStream.WriteAsync(new byte[] { 0x05, 0x05, 0x00, 0x01, 0, 0, 0, 0, 0, 0 });
        });

        using var client = new TcpClient();
        await client.ConnectAsync(IPAddress.Loopback, port);
        using var stream = client.GetStream();

        var ex = await Assert.ThrowsAsync<IOException>(() =>
            Socks5Connector.NegotiateAsync(stream, "target.example.com", 6112, null, null, CancellationToken.None));
        Assert.Contains("connection refused", ex.Message);
        await serverTask;
    }

    [Fact]
    public async Task HttpConnect_SuccessfulConnect_CompletesNegotiation()
    {
        var (listener, port) = StartFakeProxy();
        using var _ = listener;

        var serverTask = Task.Run(async () =>
        {
            using var proxyClient = await listener.AcceptTcpClientAsync();
            using var proxyStream = proxyClient.GetStream();
            using var reader = new StreamReader(proxyStream, System.Text.Encoding.ASCII, leaveOpen: true);

            var requestLine = await reader.ReadLineAsync();
            Assert.Equal("CONNECT target.example.com:6112 HTTP/1.1", requestLine);
            string? line;
            do
            {
                line = await reader.ReadLineAsync();
            } while (!string.IsNullOrEmpty(line));

            var responseBytes = System.Text.Encoding.ASCII.GetBytes("HTTP/1.1 200 Connection Established\r\n\r\n");
            await proxyStream.WriteAsync(responseBytes);
        });

        using var client = new TcpClient();
        await client.ConnectAsync(IPAddress.Loopback, port);
        using var stream = client.GetStream();

        await HttpConnectProxyConnector.NegotiateAsync(stream, "target.example.com", 6112, null, null, CancellationToken.None);
        await serverTask;
    }

    [Fact]
    public async Task HttpConnect_ProxyRefuses_ThrowsWithStatusLine()
    {
        var (listener, port) = StartFakeProxy();
        using var _ = listener;

        var serverTask = Task.Run(async () =>
        {
            using var proxyClient = await listener.AcceptTcpClientAsync();
            using var proxyStream = proxyClient.GetStream();
            using var reader = new StreamReader(proxyStream, System.Text.Encoding.ASCII, leaveOpen: true);
            await reader.ReadLineAsync();
            string? line;
            do
            {
                line = await reader.ReadLineAsync();
            } while (!string.IsNullOrEmpty(line));

            var responseBytes = System.Text.Encoding.ASCII.GetBytes("HTTP/1.1 407 Proxy Authentication Required\r\n\r\n");
            await proxyStream.WriteAsync(responseBytes);
        });

        using var client = new TcpClient();
        await client.ConnectAsync(IPAddress.Loopback, port);
        using var stream = client.GetStream();

        var ex = await Assert.ThrowsAsync<IOException>(() =>
            HttpConnectProxyConnector.NegotiateAsync(stream, "target.example.com", 6112, null, null, CancellationToken.None));
        Assert.Contains("407", ex.Message);
        await serverTask;
    }

    private static async Task<byte[]> ReadExactAsync(NetworkStream stream, int count)
    {
        var buffer = new byte[count];
        var offset = 0;
        while (offset < count)
        {
            var read = await stream.ReadAsync(buffer.AsMemory(offset, count - offset));
            if (read == 0)
            {
                throw new IOException("Unexpected end of stream in test.");
            }

            offset += read;
        }

        return buffer;
    }
}
