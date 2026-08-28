using System.Net.Sockets;

namespace Invigoration.Core.Networking;

/// <summary>
/// Async, length-prefixed TCP client. Cross-platform (System.Net.Sockets, no
/// platform-specific APIs) replacement for the VB6 Winsock control pattern of
/// buffering DataArrival chunks and splitting them into whole packets.
/// Subclasses supply the framing rule for their protocol (BNCS/BNLS/realm all
/// use slightly different header layouts).
/// </summary>
public abstract class FramedTcpClient : IAsyncDisposable
{
    private readonly List<byte> _receiveBuffer = [];
    private TcpClient? _client;
    private NetworkStream? _stream;
    private CancellationTokenSource? _receiveCts;

    public event Action? Connected;
    public event Action<Exception?>? Disconnected;
    public event Action<byte[]>? PacketReceived;

    public bool IsConnected => _client?.Connected ?? false;

    /// <summary>
    /// Given the bytes buffered so far, returns the total length of the next
    /// complete frame (header included), or null if more data is needed.
    /// </summary>
    protected abstract int? TryGetFrameLength(IReadOnlyList<byte> buffer);

    /// <summary>Connects directly to host:port, or — when <paramref name="proxy"/> is given — connects to the proxy first and tunnels to host:port through it (SOCKS5 or HTTP CONNECT), so the target server sees the proxy's IP instead of this machine's.</summary>
    public async Task ConnectAsync(string host, int port, CancellationToken cancellationToken = default, ProxyOptions? proxy = null)
    {
        Close();

        var client = new TcpClient();
        var connectHost = proxy?.Host ?? host;
        var connectPort = proxy?.Port ?? port;
        await client.ConnectAsync(connectHost, connectPort, cancellationToken).ConfigureAwait(false);

        var stream = client.GetStream();
        if (proxy is not null)
        {
            try
            {
                switch (proxy.Protocol)
                {
                    case ProxyProtocol.Socks5:
                        await Socks5Connector.NegotiateAsync(stream, host, port, proxy.Username, proxy.Password, cancellationToken)
                            .ConfigureAwait(false);
                        break;

                    case ProxyProtocol.Http:
                        await HttpConnectProxyConnector.NegotiateAsync(stream, host, port, proxy.Username, proxy.Password, cancellationToken)
                            .ConfigureAwait(false);
                        break;
                }
            }
            catch
            {
                client.Dispose();
                throw;
            }
        }

        _client = client;
        _stream = stream;
        _receiveBuffer.Clear();

        Connected?.Invoke();

        _receiveCts = new CancellationTokenSource();
        _ = ReceiveLoopAsync(_receiveCts.Token);
    }

    public async Task SendAsync(byte[] packet, CancellationToken cancellationToken = default)
    {
        var stream = _stream;
        if (stream is null)
        {
            return;
        }

        try
        {
            await stream.WriteAsync(packet, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is IOException or ObjectDisposedException or SocketException)
        {
            // The socket died between the null-check above and this write (or was already dead
            // but Close() hadn't run yet) — treat it exactly like a failed read: tear the
            // connection down via Close() (which drives the receive loop to its own natural exit
            // and fires Disconnected) instead of letting the exception escape into caller code.
            // Confirmed necessary live: an uncaught exception here crashed the whole app when a
            // Hotline session tried to send a chat message on an already-broken connection — a
            // real defect in this shared base class, not something specific to Hotline.
            Close();
        }
    }

    public void Close()
    {
        _receiveCts?.Cancel();
        _receiveCts = null;
        _stream?.Dispose();
        _stream = null;
        _client?.Dispose();
        _client = null;
    }

    private async Task ReceiveLoopAsync(CancellationToken cancellationToken)
    {
        var readBuffer = new byte[8192];
        Exception? failure = null;

        try
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                var stream = _stream;
                if (stream is null)
                {
                    break;
                }

                var bytesRead = await stream.ReadAsync(readBuffer, cancellationToken).ConfigureAwait(false);
                if (bytesRead == 0)
                {
                    break; // remote closed the connection
                }

                _receiveBuffer.AddRange(readBuffer.AsSpan(0, bytesRead).ToArray());

                while (true)
                {
                    var frameLength = TryGetFrameLength(_receiveBuffer);
                    if (frameLength is null || _receiveBuffer.Count < frameLength.Value)
                    {
                        break;
                    }

                    var frame = _receiveBuffer.GetRange(0, frameLength.Value).ToArray();
                    _receiveBuffer.RemoveRange(0, frameLength.Value);
                    PacketReceived?.Invoke(frame);
                }
            }
        }
        catch (OperationCanceledException)
        {
            // Close() was called; not a failure.
        }
        catch (Exception ex)
        {
            failure = ex;
        }

        Disconnected?.Invoke(failure);
    }

    public ValueTask DisposeAsync()
    {
        Close();
        return ValueTask.CompletedTask;
    }
}
