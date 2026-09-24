using System.Net.Security;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;

namespace Invigoration.Scr.Classic;

/// <summary>
/// A minimal RFC 6455 client over TLS. It exists because the classic connection's scrambling
/// seed is folded from the handshake key (see ClassicEnvelope.SeedFromWebSocketKey), and .NET's
/// ClientWebSocket keeps that key private.
/// </summary>
public sealed class ClassicWebSocket : IAsyncDisposable
{
    private readonly TcpClient _tcp = new() { NoDelay = true };
    private Stream _stream = Stream.Null;
    private readonly SemaphoreSlim _sendLock = new(1, 1);

    public string Key { get; } = Convert.ToBase64String(RandomNumberGenerator.GetBytes(16));

    public string? SelectedProtocol { get; private set; }

    /// <summary>The close code and reason the server sent, if it closed with a close frame.</summary>
    public int? CloseCode { get; private set; }

    public string? CloseReason { get; private set; }

    public async Task ConnectAsync(Uri uri, IReadOnlyList<string> protocols, CancellationToken token)
    {
        var port = uri.IsDefaultPort ? (uri.Scheme == "wss" ? 443 : 80) : uri.Port;
        await _tcp.ConnectAsync(uri.Host, port, token);
        Stream stream = _tcp.GetStream();
        if (uri.Scheme == "wss")
        {
            var ssl = new SslStream(stream);
            await ssl.AuthenticateAsClientAsync(uri.Host);
            stream = ssl;
        }

        _stream = stream;
        var request = new StringBuilder()
            .Append($"GET {uri.PathAndQuery} HTTP/1.1\r\n")
            .Append($"Host: {uri.Host}{(uri.IsDefaultPort ? "" : $":{uri.Port}")}\r\n")
            .Append("Upgrade: websocket\r\nConnection: Upgrade\r\n")
            .Append($"Sec-WebSocket-Key: {Key}\r\nSec-WebSocket-Version: 13\r\n");
        if (protocols.Count > 0)
        {
            request.Append($"Sec-WebSocket-Protocol: {string.Join(", ", protocols)}\r\n");
        }

        request.Append("\r\n");
        await _stream.WriteAsync(Encoding.ASCII.GetBytes(request.ToString()), token);

        var response = await ReadHeadersAsync(token);
        if (!response.StartsWith("HTTP/1.1 101", StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"WebSocket upgrade refused: {response.Split("\r\n")[0]}");
        }

        var expected = Convert.ToBase64String(SHA1.HashData(Encoding.ASCII.GetBytes(Key + "258EAFA5-E914-47DA-95CA-C5AB0DC85B11")));
        var headers = response.Split("\r\n").Skip(1).Select(l => l.Split(':', 2)).Where(p => p.Length == 2)
            .ToDictionary(p => p[0].Trim().ToLowerInvariant(), p => p[1].Trim());
        if (!headers.TryGetValue("sec-websocket-accept", out var accept) || accept != expected)
        {
            throw new InvalidOperationException("WebSocket upgrade had a wrong Sec-WebSocket-Accept.");
        }

        headers.TryGetValue("sec-websocket-protocol", out var selected);
        SelectedProtocol = selected;
    }

    private async Task<string> ReadHeadersAsync(CancellationToken token)
    {
        var bytes = new List<byte>();
        var one = new byte[1];
        while (!(bytes.Count >= 4 && bytes[^4] == '\r' && bytes[^3] == '\n' && bytes[^2] == '\r' && bytes[^1] == '\n'))
        {
            if (await _stream.ReadAsync(one, token) == 0)
            {
                throw new IOException("Connection closed during the WebSocket handshake.");
            }

            bytes.Add(one[0]);
        }

        return Encoding.ASCII.GetString(bytes.ToArray());
    }

    /// <summary>One whole message: (isBinary, payload), or null when the socket closed. Pings are answered here.</summary>
    public async Task<(bool Binary, byte[] Data)?> ReceiveAsync(CancellationToken token)
    {
        var message = new MemoryStream();
        var binary = false;
        while (true)
        {
            var head = await ReadExactAsync(2, token);
            if (head is null)
            {
                return null;
            }

            var fin = (head[0] & 0x80) != 0;
            var opcode = head[0] & 0x0F;
            long length = head[1] & 0x7F;
            if (length == 126)
            {
                var ext = await ReadExactAsync(2, token) ?? throw new IOException("Closed mid-frame.");
                length = (ext[0] << 8) | ext[1];
            }
            else if (length == 127)
            {
                var ext = await ReadExactAsync(8, token) ?? throw new IOException("Closed mid-frame.");
                length = (long)System.Buffers.Binary.BinaryPrimitives.ReadUInt64BigEndian(ext);
            }

            var payload = length == 0 ? [] : await ReadExactAsync((int)length, token) ?? throw new IOException("Closed mid-frame.");
            switch (opcode)
            {
                case 0x8:
                    CloseCode = payload.Length >= 2 ? (payload[0] << 8) | payload[1] : null;
                    CloseReason = payload.Length > 2 ? Encoding.UTF8.GetString(payload, 2, payload.Length - 2) : "";
                    return null;
                case 0x9:
                    await SendFrameAsync(0xA, payload, token);
                    continue;
                case 0xA:
                    continue;
                case 0x1 or 0x2:
                    binary = opcode == 0x2;
                    break;
            }

            message.Write(payload);
            if (fin)
            {
                return (binary, message.ToArray());
            }
        }
    }

    public Task SendAsync(byte[] data, bool binary, CancellationToken token) => SendFrameAsync(binary ? 0x2 : 0x1, data, token);

    private async Task SendFrameAsync(int opcode, byte[] payload, CancellationToken token)
    {
        var frame = new MemoryStream();
        frame.WriteByte((byte)(0x80 | opcode));
        if (payload.Length < 126)
        {
            frame.WriteByte((byte)(0x80 | payload.Length));
        }
        else if (payload.Length <= ushort.MaxValue)
        {
            frame.WriteByte(0x80 | 126);
            frame.WriteByte((byte)(payload.Length >> 8));
            frame.WriteByte((byte)payload.Length);
        }
        else
        {
            frame.WriteByte(0x80 | 127);
            var ext = new byte[8];
            System.Buffers.Binary.BinaryPrimitives.WriteUInt64BigEndian(ext, (ulong)payload.Length);
            frame.Write(ext);
        }

        var mask = RandomNumberGenerator.GetBytes(4);
        frame.Write(mask);
        for (var i = 0; i < payload.Length; i++)
        {
            frame.WriteByte((byte)(payload[i] ^ mask[i % 4]));
        }

        await _sendLock.WaitAsync(token);
        try
        {
            await _stream.WriteAsync(frame.ToArray(), token);
            await _stream.FlushAsync(token);
        }
        finally
        {
            _sendLock.Release();
        }
    }

    private async Task<byte[]?> ReadExactAsync(int count, CancellationToken token)
    {
        var buffer = new byte[count];
        var read = 0;
        while (read < count)
        {
            var n = await _stream.ReadAsync(buffer.AsMemory(read), token);
            if (n == 0)
            {
                return null;
            }

            read += n;
        }

        return buffer;
    }

    public async ValueTask DisposeAsync()
    {
        try
        {
            await SendFrameAsync(0x8, [], CancellationToken.None);
        }
        catch
        {
            // Already gone.
        }

        _tcp.Dispose();
    }
}
