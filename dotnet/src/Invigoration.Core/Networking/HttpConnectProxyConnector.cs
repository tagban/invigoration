using System.Net.Sockets;
using System.Text;

namespace Invigoration.Core.Networking;

/// <summary>Tunnels a raw TCP connection through an HTTP proxy's CONNECT method (RFC 7231 §4.3.6) — simpler than SOCKS5, widely supported by corporate/VPN-style proxies.</summary>
public static class HttpConnectProxyConnector
{
    public static async Task NegotiateAsync(NetworkStream stream, string targetHost, int targetPort, string? username, string? password, CancellationToken cancellationToken)
    {
        var requestBytes = Encoding.ASCII.GetBytes(BuildConnectRequest(targetHost, targetPort, username, password));
        await stream.WriteAsync(requestBytes, cancellationToken).ConfigureAwait(false);

        var statusLine = await ReadStatusLineAsync(stream, cancellationToken).ConfigureAwait(false);
        if (!IsSuccessStatusLine(statusLine))
        {
            throw new IOException($"HTTP proxy refused the CONNECT to {targetHost}:{targetPort}: \"{statusLine}\".");
        }
    }

    public static string BuildConnectRequest(string host, int port, string? username, string? password)
    {
        var hostPort = $"{host}:{port}";
        var request = $"CONNECT {hostPort} HTTP/1.1\r\nHost: {hostPort}\r\n";
        if (!string.IsNullOrEmpty(username))
        {
            var credentials = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{username}:{password}"));
            request += $"Proxy-Authorization: Basic {credentials}\r\n";
        }

        return request + "\r\n";
    }

    public static bool IsSuccessStatusLine(string statusLine)
    {
        // "HTTP/1.1 200 Connection Established" (exact reason phrase varies by proxy).
        var parts = statusLine.Split(' ', 3);
        return parts.Length >= 2 && parts[1].Length == 3 && parts[1][0] == '2';
    }

    /// <summary>Reads bytes one at a time until the first "\r\n" (the response's status line), discarding the rest of the headers up to the blank line that ends them — this bot doesn't need any of them.</summary>
    private static async Task<string> ReadStatusLineAsync(NetworkStream stream, CancellationToken cancellationToken)
    {
        var line = await ReadLineAsync(stream, cancellationToken).ConfigureAwait(false);

        // Drain the remaining headers up to the blank line terminating them.
        string next;
        do
        {
            next = await ReadLineAsync(stream, cancellationToken).ConfigureAwait(false);
        } while (next.Length > 0);

        return line;
    }

    private static async Task<string> ReadLineAsync(NetworkStream stream, CancellationToken cancellationToken)
    {
        var bytes = new List<byte>();
        var single = new byte[1];
        while (true)
        {
            var read = await stream.ReadAsync(single, cancellationToken).ConfigureAwait(false);
            if (read == 0)
            {
                throw new IOException("HTTP proxy closed the connection unexpectedly during CONNECT negotiation.");
            }

            if (single[0] == '\n')
            {
                if (bytes.Count > 0 && bytes[^1] == '\r')
                {
                    bytes.RemoveAt(bytes.Count - 1);
                }

                break;
            }

            bytes.Add(single[0]);
        }

        return Encoding.ASCII.GetString(bytes.ToArray());
    }
}
