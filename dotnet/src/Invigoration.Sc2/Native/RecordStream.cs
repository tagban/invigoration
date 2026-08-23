using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

/// <summary>
/// Wraps a duplex <see cref="Stream"/> (a live TCP/TLS socket in production,
/// a MemoryStream in tests) with SC2's native transport framing: outbound
/// records are RC4-encrypted once <see cref="EnableEncryption"/> has been
/// called (plaintext before that, matching the Resume handshake), and
/// inbound bytes are decrypted immediately as they arrive — RC4 is a
/// stateful stream cipher, so bytes must be fed through it in the exact
/// order they were received, independent of how the OS chooses to
/// segment the TCP stream. Mirrors core/src/native/stream.rs's RecordStream.
///
/// There is no length prefix on the wire: a record's end is only knowable by
/// decoding it (see the "Sunken" walkthrough at
/// https://superioritybot.com/PROTOCOL#packets). Without upstream's full
/// schema table, this class can only hand off routes the caller supplies a
/// decoder for via <see cref="TryDecodeRecord{T}"/> — an unrecognized route
/// cannot be skipped, since its bit width isn't known, and would desync the
/// stream. This is an inherent limitation of the schema-free port, not a
/// simplification made here.
/// </summary>
public sealed class RecordStream : IDisposable
{
    private readonly Stream _stream;
    private readonly List<byte> _inbound = [];
    private Rc4State? _inboundCipher;
    private Rc4State? _outboundCipher;

    public RecordStream(Stream stream) => _stream = stream;

    public bool IsProtected => _outboundCipher is not null;

    /// <summary>Switches both directions to RC4 from this point on. Call once, immediately after sending the plaintext Conn/5 EnableEncryption record.</summary>
    public void EnableEncryption(Rc4State inboundCipher, Rc4State outboundCipher)
    {
        _inboundCipher = inboundCipher;
        _outboundCipher = outboundCipher;
    }

    public async Task SendAsync(byte[] record, CancellationToken cancellationToken = default)
    {
        var payload = _outboundCipher is null ? record : _outboundCipher.Apply(record);
        await _stream.WriteAsync(payload, cancellationToken).ConfigureAwait(false);
        await _stream.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Reads whatever is currently available from the socket, decrypts it in place if encryption is active, and appends it to the pending buffer. Returns false on end-of-stream.</summary>
    public async Task<bool> FillAsync(CancellationToken cancellationToken = default)
    {
        var chunk = new byte[4096];
        var read = await _stream.ReadAsync(chunk, cancellationToken).ConfigureAwait(false);
        if (read == 0)
        {
            return false;
        }

        var received = chunk.AsSpan(0, read).ToArray();
        _inboundCipher?.ApplyInPlace(received);
        _inbound.AddRange(received);
        return true;
    }

    /// <summary>
    /// Attempts to decode exactly one record from the front of the pending
    /// buffer: decodes the routing header, then hands the positioned reader
    /// to <paramref name="decode"/> for the payload. On success, advances
    /// past the record (aligning to the next byte boundary) and returns
    /// true. On a buffer that doesn't yet hold a complete record, leaves the
    /// buffer untouched and returns false — call <see cref="FillAsync"/>
    /// again and retry. Exceptions from <paramref name="decode"/> other than
    /// the ones used internally to signal "not enough data yet" propagate
    /// normally (e.g. an unrecognized choice selector is a real protocol
    /// error, not a buffering issue).
    /// </summary>
    public bool TryDecodeRecord<T>(Func<byte, byte?, BitReader, T> decode, out T? record)
    {
        var snapshot = _inbound.ToArray();
        var reader = new BitReader(snapshot);
        try
        {
            var routing = RoutingHeader.Decode(reader);
            var value = decode(routing.CommandId, routing.ServiceSlot, reader);
            reader.Align();
            var consumedBytes = reader.Position / 8;
            _inbound.RemoveRange(0, consumedBytes);
            record = value;
            return true;
        }
        catch (Exception ex) when (IsBufferUnderrun(ex))
        {
            record = default;
            return false;
        }
    }

    private static bool IsBufferUnderrun(Exception ex) =>
        ex is IndexOutOfRangeException or ArgumentOutOfRangeException or ArgumentException;

    /// <summary>Diagnostic-only: hex dump of the buffered-but-not-yet-consumed bytes, for capturing a real record this project doesn't have a decoder for yet.</summary>
    public string PendingHex() => Convert.ToHexString(_inbound.ToArray());

    public void Dispose() => _stream.Dispose();
}
