namespace Invigoration.Sc2.Wire;

/// <summary>Reads SC2's "BSN" bit-packed wire format. Mirrors <see cref="BitWriter"/> exactly — see its remarks for the two coexisting bit orders.</summary>
public sealed class BitReader
{
    private readonly byte[] _data;
    private int _position;

    public BitReader(byte[] data, int startPosition = 0)
    {
        _data = data;
        _position = startPosition;
    }

    public int Position => _position;

    public int RemainingBits => (_data.Length * 8) - _position;

    /// <summary>Reads <paramref name="width"/> bits and reassembles them MSB-first into the returned value.</summary>
    public ulong Read(int width)
    {
        ulong value = 0;
        var remaining = width;
        while (remaining > 0)
        {
            var byteIndex = _position / 8;
            var bitOffset = _position % 8;
            var take = Math.Min(remaining, 8 - bitOffset);
            var mask = (byte)((1 << take) - 1);
            var chunk = (byte)((_data[byteIndex] >> bitOffset) & mask);

            value = (value << take) | chunk;
            _position += take;
            remaining -= take;
        }

        return value;
    }

    /// <summary>Extracts <paramref name="bitCount"/> bits as a plain LSB-first bit copy (no value reordering) — the inverse of <see cref="BitWriter.WriteRaw"/>.</summary>
    public byte[] ReadRaw(int bitCount)
    {
        var writer = new BitWriter();
        for (var i = 0; i < bitCount; i++)
        {
            writer.Write(Read(1), 1);
        }

        return writer.ToBytes();
    }

    /// <summary>Reads raw bytes. If <paramref name="aligned"/>, skips to the next byte boundary first (via <see cref="Align"/>), mirroring <see cref="BitWriter.WriteBytes"/>. Otherwise falls back to <see cref="ReadRaw"/>.</summary>
    public byte[] ReadBytes(int count, bool aligned)
    {
        if (aligned)
        {
            Align();
            var byteIndex = _position / 8;
            var result = new byte[count];
            Array.Copy(_data, byteIndex, result, 0, count);
            _position += count * 8;
            return result;
        }

        return ReadRaw(count * 8);
    }

    /// <summary>Skips to the next byte boundary. Returns the number of bits skipped.</summary>
    public int Align()
    {
        var skipped = (-_position) & 7;
        _position += skipped;
        return skipped;
    }
}
