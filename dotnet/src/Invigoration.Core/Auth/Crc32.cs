namespace Invigoration.Core.Auth;

/// <summary>
/// Standard reflected CRC-32 (poly 0xEDB88320, init 0xFFFFFFFF, final XOR),
/// as used by BNLS's checksum challenge-response. Direct port of the table
/// generation and update loop from modBNLS.bas.
/// </summary>
public static class Crc32
{
    private const uint Polynomial = 0xEDB88320;

    private static readonly uint[] Table = BuildTable();

    private static uint[] BuildTable()
    {
        var table = new uint[256];
        for (uint i = 0; i < 256; i++)
        {
            var value = i;
            for (var bit = 0; bit < 8; bit++)
            {
                value = (value & 1) != 0
                    ? (value >> 1) ^ Polynomial
                    : value >> 1;
            }

            table[i] = value;
        }

        return table;
    }

    public static uint Compute(ReadOnlySpan<byte> data)
    {
        var crc = 0xFFFFFFFFu;
        foreach (var b in data)
        {
            var index = (byte)(b ^ (crc & 0xFF));
            crc = (crc >> 8) ^ Table[index];
        }

        return ~crc;
    }
}
