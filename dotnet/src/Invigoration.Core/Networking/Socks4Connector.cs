using System.Net;
using System.Net.Sockets;
using System.Text;

namespace Invigoration.Core.Networking;

/// <summary>
/// Minimal SOCKS4 / SOCKS4a client (CONNECT only). SOCKS4 carries an IPv4 address; for a host
/// name it uses SOCKS4a (address 0.0.0.1 and the name after the user ID), so the proxy looks the
/// name up. SOCKS4 has no password: a username goes in the user-ID field. Wire building and
/// parsing are static so they're testable without a proxy, as in <see cref="Socks5Connector"/>.
/// </summary>
public static class Socks4Connector
{
    private const byte Version = 0x04;
    private const byte CommandConnect = 0x01;
    private const byte Granted = 0x5A;

    public static async Task NegotiateAsync(NetworkStream stream, string targetHost, int targetPort, string? username, CancellationToken cancellationToken)
    {
        await stream.WriteAsync(BuildConnectRequest(targetHost, targetPort, username), cancellationToken).ConfigureAwait(false);
        var reply = new byte[8];
        var read = 0;
        while (read < reply.Length)
        {
            var n = await stream.ReadAsync(reply.AsMemory(read), cancellationToken).ConfigureAwait(false);
            if (n == 0)
            {
                throw new IOException("SOCKS4 proxy closed the connection during the handshake.");
            }

            read += n;
        }

        if (!IsGranted(reply))
        {
            throw new IOException($"SOCKS4 proxy refused the connection to {targetHost}:{targetPort} ({DescribeReplyCode(reply[1])}).");
        }
    }

    /// <summary>VN 4, CD 1, port (big-endian), IPv4 address, user ID, 0; for SOCKS4a, address 0.0.0.1 and the host name, 0, after it.</summary>
    public static byte[] BuildConnectRequest(string host, int port, string? userId)
    {
        var user = Encoding.ASCII.GetBytes(userId ?? "");
        var isIpv4 = IPAddress.TryParse(host, out var ip) && ip.AddressFamily == AddressFamily.InterNetwork;
        var name = isIpv4 ? [] : Encoding.ASCII.GetBytes(host);
        var buffer = new List<byte>(9 + user.Length + name.Length + 1)
        {
            Version,
            CommandConnect,
            (byte)(port >> 8),
            (byte)port,
        };
        buffer.AddRange(isIpv4 ? ip!.GetAddressBytes() : [0, 0, 0, 1]);
        buffer.AddRange(user);
        buffer.Add(0);
        if (!isIpv4)
        {
            buffer.AddRange(name);
            buffer.Add(0);
        }

        return [.. buffer];
    }

    public static bool IsGranted(byte[] reply) => reply.Length >= 2 && reply[1] == Granted;

    private static string DescribeReplyCode(byte code) => code switch
    {
        0x5B => "request rejected or failed",
        0x5C => "the proxy couldn't reach identd on this machine",
        0x5D => "identd reported a different user ID",
        _ => $"unknown reply 0x{code:X2}",
    };
}
