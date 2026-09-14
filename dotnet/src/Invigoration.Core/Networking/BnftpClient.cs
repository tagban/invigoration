using System.Net.Sockets;
using System.Text;

namespace Invigoration.Core.Networking;

/// <summary>A file fetched over BNFTP: the server's name for it, its file time, and its bytes.</summary>
public sealed record BnftpFile(string Name, long FileTime, byte[] Data);

/// <summary>
/// Battle.net File Transfer Protocol, version 1 — how classic clients download icons.bni and ad
/// banners from the Battle.net server itself, on the same port as chat. One request per
/// connection: send protocol byte 0x02 and a request naming the file, get back a header with its
/// size and time, then the file. Per bnetdocs.org "File Transfer Protocol Version 1".
/// </summary>
public static class BnftpClient
{
    public const int DefaultPort = 6112;

    /// <summary>Refuse anything claiming to be bigger than this — a server could announce any size.</summary>
    public const int MaxFileBytes = 64 * 1024 * 1024;

    private const ushort ProtocolVersion = 0x100;

    /// <summary>
    /// Downloads <paramref name="fileName"/>, or returns null when the server doesn't have it (it
    /// closes the connection without replying). Identifies as StarCraft on Intel x86 — the server
    /// only needs a known platform and product to serve a plain file.
    /// </summary>
    /// <exception cref="IOException">The reply was malformed, truncated, or too large.</exception>
    public static async Task<BnftpFile?> DownloadAsync(string host, int port, string fileName, TimeSpan timeout, CancellationToken cancellationToken = default)
    {
        using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        cts.CancelAfter(timeout);
        var ct = cts.Token;

        using var client = new TcpClient();
        await client.ConnectAsync(host, port, ct).ConfigureAwait(false);
        await using var stream = client.GetStream();

        await stream.WriteAsync(BuildRequest(fileName), ct).ConfigureAwait(false);

        var lengthBytes = new byte[2];
        if (!await ReadExactlyOrEndAsync(stream, lengthBytes, ct).ConfigureAwait(false))
        {
            return null;
        }

        var headerLength = BitConverter.ToUInt16(lengthBytes);
        if (headerLength < 2 + 2 + 4 + 4 + 4 + 8 + 1)
        {
            throw new IOException($"BNFTP header too short ({headerLength} bytes).");
        }

        var header = new byte[headerLength - 2];
        if (!await ReadExactlyOrEndAsync(stream, header, ct).ConfigureAwait(false))
        {
            throw new IOException("BNFTP header was cut off.");
        }

        var size = BitConverter.ToUInt32(header, 2);
        var fileTime = BitConverter.ToInt64(header, 14);
        var nameEnd = Array.IndexOf(header, (byte)0, 22);
        var name = Encoding.ASCII.GetString(header, 22, (nameEnd < 0 ? header.Length : nameEnd) - 22);
        if (size > MaxFileBytes)
        {
            throw new IOException($"BNFTP file \"{name}\" claims {size} bytes, more than the {MaxFileBytes}-byte limit.");
        }

        var data = new byte[size];
        if (!await ReadExactlyOrEndAsync(stream, data, ct).ConfigureAwait(false))
        {
            throw new IOException($"BNFTP file \"{name}\" was cut off before its {size} bytes arrived.");
        }

        return new BnftpFile(name, fileTime, data);
    }

    /// <summary>The bytes of a version 1 request for <paramref name="fileName"/>, protocol selector included.</summary>
    public static byte[] BuildRequest(string fileName)
    {
        var name = Encoding.ASCII.GetBytes(fileName);
        var body = new List<byte>();
        body.AddRange(BitConverter.GetBytes(ProtocolVersion));
        body.AddRange("68XI"u8.ToArray()); // platform "IX86", as the little-endian DWORD the wire carries
        body.AddRange("RATS"u8.ToArray()); // product "STAR", likewise
        body.AddRange(BitConverter.GetBytes(0u)); // banner id
        body.AddRange(BitConverter.GetBytes(0u)); // banner file extension
        body.AddRange(BitConverter.GetBytes(0u)); // start at the beginning
        body.AddRange(BitConverter.GetBytes(0L)); // no local copy's file time
        body.AddRange(name);
        body.Add(0);

        var request = new List<byte> { 0x02 }; // protocol selector: BNFTP
        request.AddRange(BitConverter.GetBytes((ushort)(body.Count + 2)));
        request.AddRange(body);
        return request.ToArray();
    }

    /// <summary>False if the connection closes before any byte arrives; throws if it closes partway.</summary>
    private static async Task<bool> ReadExactlyOrEndAsync(NetworkStream stream, byte[] buffer, CancellationToken ct)
    {
        var read = 0;
        while (read < buffer.Length)
        {
            var n = await stream.ReadAsync(buffer.AsMemory(read), ct).ConfigureAwait(false);
            if (n == 0)
            {
                if (read == 0)
                {
                    return false;
                }

                throw new IOException("Connection closed partway through a BNFTP reply.");
            }

            read += n;
        }

        return true;
    }
}
