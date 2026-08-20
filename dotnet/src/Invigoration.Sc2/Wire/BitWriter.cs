namespace Invigoration.Sc2.Wire;

/// <summary>
/// Writes SC2's "BSN" bit-packed wire format. Ported from
/// ncarrillo/superiority's core/src/bsn/bits.rs — verified against that
/// crate's own unit-test vectors (see BitWriterTests). Two distinct bit
/// orders coexist here and must not be confused:
///
/// <list type="bullet">
/// <item><see cref="Write"/> consumes an integer value MSB-first, but packs
/// each chunk into the current byte LSB-first (low bits of the byte filled
/// before high bits) — this is what every integer/enum/range field uses.</item>
/// <item><see cref="WriteRaw"/> is a plain LSB-first bit-for-bit copy with no
/// value reinterpretation — used for bit arrays and pre-encoded byte blobs
/// spliced in at a non-byte-aligned position.</item>
/// </list>
/// </summary>
public sealed class BitWriter
{
    private byte[] _data = new byte[16];
    private int _position;

    public int Position => _position;

    /// <summary>Writes the low <paramref name="width"/> bits of <paramref name="value"/>, MSB-first.</summary>
    public void Write(ulong value, int width)
    {
        var remaining = width;
        while (remaining > 0)
        {
            var byteIndex = _position / 8;
            var bitOffset = _position % 8;
            var take = Math.Min(remaining, 8 - bitOffset);
            var shift = remaining - take;
            var mask = (byte)((1 << take) - 1);
            var chunk = (byte)((value >> shift) & mask);

            EnsureCapacity(byteIndex + 1);
            _data[byteIndex] &= (byte)~(mask << bitOffset);
            _data[byteIndex] |= (byte)(chunk << bitOffset);

            _position += take;
            remaining -= take;
        }
    }

    /// <summary>Copies <paramref name="bitCount"/> bits from <paramref name="source"/> (LSB-first, no reordering) starting at bit 0.</summary>
    public void WriteRaw(ReadOnlySpan<byte> source, int bitCount)
    {
        for (var sourcePosition = 0; sourcePosition < bitCount; sourcePosition++)
        {
            var bit = (source[sourcePosition / 8] >> (sourcePosition & 7)) & 1;
            var byteIndex = _position / 8;
            EnsureCapacity(byteIndex + 1);
            if (bit != 0)
            {
                _data[byteIndex] |= (byte)(1 << (_position & 7));
            }

            _position++;
        }
    }

    /// <summary>
    /// Writes raw bytes. If <paramref name="aligned"/>, pads with zero bits to the next byte
    /// boundary first (via <see cref="Align"/>) and then blits directly — this is NOT "copy
    /// directly if already aligned, else fall back to a raw bit copy"; alignment is forced.
    /// Verified against the toon_select golden vector, where 6 padding bits appear between the
    /// (11-bit routing header + 7-bit length) and the byte-aligned name. Otherwise, falls back to
    /// <see cref="WriteRaw"/> for a true unaligned blit.
    /// </summary>
    public void WriteBytes(ReadOnlySpan<byte> bytes, bool aligned)
    {
        if (aligned)
        {
            Align();
            var byteIndex = _position / 8;
            EnsureCapacity(byteIndex + bytes.Length);
            bytes.CopyTo(_data.AsSpan(byteIndex));
            _position += bytes.Length * 8;
            return;
        }

        WriteRaw(bytes, bytes.Length * 8);
    }

    /// <summary>Pads with zero bits to the next byte boundary. Returns the number of bits skipped.</summary>
    public int Align()
    {
        var skipped = (-_position) & 7;
        if (skipped > 0)
        {
            Write(0, skipped);
        }

        return skipped;
    }

    public byte[] ToBytes()
    {
        var length = (_position + 7) / 8;
        var result = new byte[length];
        Array.Copy(_data, result, length);
        return result;
    }

    private void EnsureCapacity(int byteCount)
    {
        if (byteCount <= _data.Length)
        {
            return;
        }

        var newLength = _data.Length;
        while (newLength < byteCount)
        {
            newLength *= 2;
        }

        Array.Resize(ref _data, newLength);
    }
}
