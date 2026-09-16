using System.Buffers.Binary;
using System.Net;
using System.Net.Sockets;
using Invigoration.Core.Hotline;

namespace Invigoration.Core.Tests;

/// <summary>
/// Auto-reconnect leans on one distinction: a server that's merely down is worth retrying, a
/// server that refused the login isn't (retrying a refusal is how an account or address gets
/// banned). That decision reads HotlineTransactionClient.LoginRefusedReason, so these pin it
/// against a stand-in server that completes the real handshake and then answers the login.
/// </summary>
public class HotlineLoginRefusalTests
{
    /// <summary>Accepts one connection, replies to the 12-byte handshake, then answers the login transaction with the given error code and text.</summary>
    private static (TcpListener Listener, Task Served) StartServer(uint errorCode, string? errorText)
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var served = Task.Run(async () =>
        {
            using var client = await listener.AcceptTcpClientAsync();
            using var stream = client.GetStream();

            var handshake = new byte[12];
            await stream.ReadExactlyAsync(handshake);

            // The 8-byte handshake reply: "TRTP" then a big-endian error code, 0 meaning accepted.
            var reply = new byte[8];
            "TRTP"u8.ToArray().CopyTo(reply, 0);
            await stream.WriteAsync(reply);

            // The client's Login transaction — read its header to learn the body length and its id,
            // so the reply can be addressed to it the way a real server would.
            var header = new byte[HotlineConstants.TransactionHeaderSize];
            await stream.ReadExactlyAsync(header);
            var id = BinaryPrimitives.ReadUInt32BigEndian(header.AsSpan(4));
            var bodyLength = (int)BinaryPrimitives.ReadUInt32BigEndian(header.AsSpan(12));
            if (bodyLength > 0)
            {
                await stream.ReadExactlyAsync(new byte[bodyLength]);
            }

            var fields = errorText is null ? [] : new[] { new HotlineField(HotlineFieldType.ErrorText, errorText) };
            await stream.WriteAsync(HotlineTransactionFrame.CreateReply(id, errorCode, fields).Encode());
            await Task.Delay(300);
        });

        return (listener, served);
    }

    [Fact]
    public async Task ARefusedLogin_ReportsWhyTheServerSaidNo()
    {
        var (listener, served) = StartServer(errorCode: 1, errorText: "You are banned from this server.");
        try
        {
            await using var client = new HotlineTransactionClient();

            var ok = await client.ConnectAndLoginAsync(
                "127.0.0.1", ((IPEndPoint)listener.LocalEndpoint).Port, "guest", "", "Guest", 414);

            Assert.False(ok);
            Assert.Equal("You are banned from this server.", client.LoginRefusedReason);
            await served;
        }
        finally
        {
            listener.Stop();
        }
    }

    /// <summary>A refusal with no explanation still has to register as a refusal, or reconnect would hammer it.</summary>
    [Fact]
    public async Task ARefusalWithNoText_StillCountsAsRefused()
    {
        var (listener, served) = StartServer(errorCode: 1, errorText: null);
        try
        {
            await using var client = new HotlineTransactionClient();

            var ok = await client.ConnectAndLoginAsync(
                "127.0.0.1", ((IPEndPoint)listener.LocalEndpoint).Port, "guest", "", "Guest", 414);

            Assert.False(ok);
            Assert.NotNull(client.LoginRefusedReason);
            await served;
        }
        finally
        {
            listener.Stop();
        }
    }

    /// <summary>
    /// The case reconnect exists for: nothing listening. It must NOT look like a refusal, or a
    /// server that's rebooting would never be retried.
    /// </summary>
    [Fact]
    public async Task AServerThatIsSimplyDown_IsNotARefusal()
    {
        // Bind and immediately release a port, so nothing is listening on a port nothing else took.
        var probe = new TcpListener(IPAddress.Loopback, 0);
        probe.Start();
        var deadPort = ((IPEndPoint)probe.LocalEndpoint).Port;
        probe.Stop();

        await using var client = new HotlineTransactionClient();

        try
        {
            var ok = await client.ConnectAndLoginAsync("127.0.0.1", deadPort, "guest", "", "Guest", 414);
            Assert.False(ok);
        }
        catch (SocketException)
        {
            // Connecting threw rather than returning false — either way it never reached a login.
        }

        Assert.Null(client.LoginRefusedReason);
    }
}
