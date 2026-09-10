using System.Collections;
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

    /// <summary>
    /// Serializes every write to the socket. NetworkStream.WriteAsync is not safe to call
    /// concurrently from multiple callers — nothing previously prevented that here, and this
    /// class is shared by every outgoing path (chat commands, BNLS/auth packets, and the
    /// keepalive SID_NULL ping, which deliberately bypasses BotEngine's chat-send flood-
    /// protection gate so it's never delayed by that — see BotEngine.RunKeepAliveLoopAsync).
    /// Two of those racing to write at the same instant could interleave their bytes on the
    /// wire, producing a malformed packet a real BNCS/PVPGN server has every reason to
    /// disconnect the client over — a plausible, previously-unnoticed cause of a connection
    /// dropping for no apparent protocol-level reason. This only ever blocks a caller for the
    /// duration of one other packet's actual write (microseconds for the small packets this
    /// protocol uses), never anything close to BotEngine's own artificial flood-protection delay,
    /// so it can't reintroduce the kind of pileup that caused the mass-join hang.
    /// </summary>
    private readonly SemaphoreSlim _writeLock = new(1, 1);

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

        await _writeLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            // Re-read _stream after acquiring the lock: Close() (from another writer's failed
            // send, or a normal disconnect) could have nulled it out while this call was waiting.
            if (_stream is not { } liveStream)
            {
                return;
            }

            await liveStream.WriteAsync(packet, cancellationToken).ConfigureAwait(false);
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
        finally
        {
            _writeLock.Release();
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

                // Extract every complete frame currently buffered against an offset *view* rather
                // than physically removing each one as it's found, then compact once at the end —
                // RemoveRange(0, frameLength) here previously shifted the entire remaining buffer
                // down on every single frame, which is O(bytes remaining) per frame. A burst that
                // lands k small frames in one read (e.g. many bots joining a channel at once, each
                // producing its own small SID_CHATEVENT) turned that into O(k * bytes) for that one
                // read — confirmed live as the cause of the client falling badly behind/appearing
                // to hang during a mass-join burst. One RemoveRange for everything consumed this
                // pass is O(bytes) total instead.
                var consumed = 0;
                while (true)
                {
                    var view = new BufferOffsetView(_receiveBuffer, consumed);
                    var frameLength = TryGetFrameLength(view);
                    if (frameLength is null || view.Count < frameLength.Value)
                    {
                        break;
                    }

                    var frame = _receiveBuffer.GetRange(consumed, frameLength.Value).ToArray();
                    consumed += frameLength.Value;
                    PacketReceived?.Invoke(frame);
                }

                if (consumed > 0)
                {
                    _receiveBuffer.RemoveRange(0, consumed);
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
        _writeLock.Dispose();
        return ValueTask.CompletedTask;
    }

    /// <summary>
    /// Zero-copy "everything from <paramref name="offset"/> onward" view of a growing
    /// <see cref="List{T}"/>, so TryGetFrameLength implementations (which only ever use
    /// <see cref="Count"/> and the indexer) can inspect not-yet-consumed bytes without the
    /// receive loop having to physically shift the buffer after every single frame.
    /// </summary>
    private sealed class BufferOffsetView(List<byte> source, int offset) : IReadOnlyList<byte>
    {
        public int Count => source.Count - offset;

        public byte this[int index] => source[offset + index];

        public IEnumerator<byte> GetEnumerator()
        {
            for (var i = offset; i < source.Count; i++)
            {
                yield return source[i];
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
