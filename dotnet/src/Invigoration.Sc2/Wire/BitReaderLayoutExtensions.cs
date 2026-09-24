namespace Invigoration.Sc2.Wire;

/// <summary>
/// Small building blocks for reading Sunken records whose content a chat client
/// doesn't need but must still consume exactly, since records carry no length.
/// Anything malformed throws <see cref="InvalidOperationException"/>, never an
/// <see cref="ArgumentException"/>: RecordStream reads the latter as "not all
/// of this record has arrived yet" and would wait forever instead of failing.
/// </summary>
public static class BitReaderLayoutExtensions
{
    /// <summary>Consumes any number of bits.</summary>
    public static void Skip(this BitReader reader, int bits)
    {
        while (bits > 0)
        {
            var take = Math.Min(bits, 64);
            reader.Read(take);
            bits -= take;
        }
    }

    /// <summary>A length of <paramref name="lengthBits"/> bits plus <paramref name="minimum"/>, then that many byte-aligned bytes.</summary>
    public static byte[] ReadBlob(this BitReader reader, int lengthBits, int minimum = 0) =>
        reader.ReadBytes((int)reader.Read(lengthBits) + minimum, aligned: true);

    /// <summary>A presence bit, then <paramref name="read"/> only if it's set.</summary>
    public static void SkipOptional(this BitReader reader, Action<BitReader> read)
    {
        if (reader.Read(1) != 0)
        {
            read(reader);
        }
    }

    /// <summary>A count read in <paramref name="bits"/> bits, rejected above <paramref name="maximum"/>.</summary>
    public static int ReadCount(this BitReader reader, int bits, int maximum, string what)
    {
        var count = (int)reader.Read(bits);
        if (count > maximum)
        {
            throw new InvalidOperationException($"{what} has too many entries ({count}, at most {maximum}).");
        }

        return count;
    }

    /// <summary>Skips ahead to byte <paramref name="totalBytes"/> of the record, counted from its first bit.</summary>
    public static void SkipToRecordByte(this BitReader reader, int recordStartBit, int totalBytes)
    {
        var endBit = recordStartBit + (totalBytes * 8);
        if (reader.Position > endBit)
        {
            throw new InvalidOperationException($"Record is already past its fixed {totalBytes}-byte length.");
        }

        reader.Skip(endBit - reader.Position);
    }
}
