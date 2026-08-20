using System.Net.Sockets;
using System.Text;

namespace Invigoration.Core.Networking;

/// <summary>
/// Minimal SOCKS5 client (RFC 1928 CONNECT command + RFC 1929 username/
/// password auth only — no BIND/UDP ASSOCIATE, nothing this bot needs).
/// Wire-format building/parsing is split into static methods that work on
/// plain byte arrays, so the tricky part (getting the protocol bytes right)
/// is unit-testable without a live proxy or even a socket.
/// </summary>
public static class Socks5Connector
{
    private const byte Version = 0x05;
    private const byte MethodNoAuth = 0x00;
    private const byte MethodUsernamePassword = 0x02;
    private const byte MethodNoAcceptable = 0xFF;
    private const byte AddressTypeDomainName = 0x03;
    private const byte CommandConnect = 0x01;

    /// <summary>Performs the full handshake over an already-connected stream to the proxy, ending with the proxy tunneling to targetHost:targetPort. Throws IOException with a descriptive message on any failure.</summary>
    public static async Task NegotiateAsync(NetworkStream stream, string targetHost, int targetPort, string? username, string? password, CancellationToken cancellationToken)
    {
        var hasCredentials = !string.IsNullOrEmpty(username);

        await stream.WriteAsync(BuildGreeting(hasCredentials), cancellationToken).ConfigureAwait(false);
        var methodReply = await ReadExactAsync(stream, 2, cancellationToken).ConfigureAwait(false);
        var method = ParseMethodSelection(methodReply);

        if (method == MethodNoAcceptable)
        {
            throw new IOException("SOCKS5 proxy rejected every offered authentication method.");
        }

        if (method == MethodUsernamePassword)
        {
            await stream.WriteAsync(BuildAuthRequest(username!, password ?? ""), cancellationToken).ConfigureAwait(false);
            var authReply = await ReadExactAsync(stream, 2, cancellationToken).ConfigureAwait(false);
            if (!IsAuthSuccess(authReply))
            {
                throw new IOException("SOCKS5 proxy rejected the username/password.");
            }
        }

        await stream.WriteAsync(BuildConnectRequest(targetHost, targetPort), cancellationToken).ConfigureAwait(false);

        // The reply's fixed header is 4 bytes (VER, REP, RSV, ATYP); what follows is
        // the bound address (length depends on ATYP — domain name needs an extra
        // length-prefixed read) then a 2-byte port. None of it is needed once read —
        // it's just how much of the reply has to be consumed before the tunnel is live.
        var replyHeader = await ReadExactAsync(stream, 4, cancellationToken).ConfigureAwait(false);
        var addressLength = replyHeader[3] switch
        {
            0x01 => 4, // IPv4
            0x04 => 16, // IPv6
            0x03 => (await ReadExactAsync(stream, 1, cancellationToken).ConfigureAwait(false))[0], // domain name: length-prefixed
            _ => throw new IOException("SOCKS5 proxy returned an unsupported address type in its reply."),
        };
        await ReadExactAsync(stream, addressLength + 2, cancellationToken).ConfigureAwait(false);

        var replyCode = replyHeader[1];
        if (replyCode != 0x00)
        {
            throw new IOException($"SOCKS5 proxy refused the connection to {targetHost}:{targetPort} ({DescribeReplyCode(replyCode)}).");
        }
    }

    public static byte[] BuildGreeting(bool hasCredentials) =>
        hasCredentials
            ? [Version, 0x02, MethodNoAuth, MethodUsernamePassword]
            : [Version, 0x01, MethodNoAuth];

    public static byte ParseMethodSelection(byte[] reply)
    {
        if (reply.Length != 2 || reply[0] != Version)
        {
            throw new IOException("SOCKS5 proxy sent an invalid method-selection reply.");
        }

        return reply[1];
    }

    public static byte[] BuildAuthRequest(string username, string password)
    {
        var userBytes = Encoding.UTF8.GetBytes(username);
        var passBytes = Encoding.UTF8.GetBytes(password);
        if (userBytes.Length > 255 || passBytes.Length > 255)
        {
            throw new ArgumentException("SOCKS5 username/password must each be 255 bytes or fewer.");
        }

        var buffer = new byte[3 + userBytes.Length + passBytes.Length];
        buffer[0] = 0x01; // auth subnegotiation version
        buffer[1] = (byte)userBytes.Length;
        userBytes.CopyTo(buffer, 2);
        buffer[2 + userBytes.Length] = (byte)passBytes.Length;
        passBytes.CopyTo(buffer, 3 + userBytes.Length);
        return buffer;
    }

    public static bool IsAuthSuccess(byte[] reply) => reply.Length == 2 && reply[1] == 0x00;

    public static byte[] BuildConnectRequest(string host, int port)
    {
        var hostBytes = Encoding.ASCII.GetBytes(host);
        if (hostBytes.Length > 255)
        {
            throw new ArgumentException("SOCKS5 target hostname must be 255 bytes or fewer.");
        }

        var buffer = new byte[7 + hostBytes.Length];
        buffer[0] = Version;
        buffer[1] = CommandConnect;
        buffer[2] = 0x00; // reserved
        buffer[3] = AddressTypeDomainName;
        buffer[4] = (byte)hostBytes.Length;
        hostBytes.CopyTo(buffer, 5);
        buffer[5 + hostBytes.Length] = (byte)(port >> 8);
        buffer[6 + hostBytes.Length] = (byte)port;
        return buffer;
    }

    private static string DescribeReplyCode(byte code) => code switch
    {
        0x01 => "general SOCKS server failure",
        0x02 => "connection not allowed by ruleset",
        0x03 => "network unreachable",
        0x04 => "host unreachable",
        0x05 => "connection refused",
        0x06 => "TTL expired",
        0x07 => "command not supported",
        0x08 => "address type not supported",
        _ => $"unknown error 0x{code:X2}",
    };

    private static async Task<byte[]> ReadExactAsync(NetworkStream stream, int count, CancellationToken cancellationToken)
    {
        var buffer = new byte[count];
        var offset = 0;
        while (offset < count)
        {
            var read = await stream.ReadAsync(buffer.AsMemory(offset, count - offset), cancellationToken).ConfigureAwait(false);
            if (read == 0)
            {
                throw new IOException("SOCKS5 proxy closed the connection unexpectedly during negotiation.");
            }

            offset += read;
        }

        return buffer;
    }
}
