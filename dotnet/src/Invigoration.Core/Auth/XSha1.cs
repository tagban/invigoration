using System.Buffers.Binary;

namespace Invigoration.Core.Auth;

/// <summary>
/// Battle.net's "Broken SHA-1" (a.k.a. X-SHA1) hashing algorithm, used by the
/// old login system to hash passwords and CD-keys instead of sending them to
/// BNLS. Structurally identical to standard SHA-1 (same round functions,
/// round constants, and finalization) except for the two quirks that earned
/// it the "broken" name: no length-suffix/0x80-bit padding at all (a final
/// partial block is simply zero-filled), and the message-schedule expansion
/// sets each extended word to a single bit — 1 &lt;&lt; (xor-combo &amp; 31) —
/// instead of rotating the xor-combo itself. Ported from Davnit/bncs.py's
/// BnetSha1 (bncs/hashing/bsha.py), cross-checked against two independent
/// published test vectors (see XSha1Tests).
/// </summary>
public static class XSha1
{
    private const int BlockSize = 64;
    private static readonly uint[] RoundConstants = [0x5A827999, 0x6ED9EBA1, 0x8F1BBCDC, 0xCA62C1D6];

    /// <summary>Hashes one or more byte segments as if they were concatenated.</summary>
    public static byte[] Hash(params byte[][] segments)
    {
        var state = new uint[] { 0x67452301, 0xEFCDAB89, 0x98BADCFE, 0x10325476, 0xC3D2E1F0 };
        var block = new byte[BlockSize];
        var position = 0;

        foreach (var segment in segments)
        {
            var offset = 0;
            while (offset < segment.Length)
            {
                var take = Math.Min(segment.Length - offset, BlockSize - position);
                Buffer.BlockCopy(segment, offset, block, position, take);
                position += take;
                offset += take;

                if (position == BlockSize)
                {
                    Transform(state, block);
                    position = 0;
                }
            }
        }

        if (position > 0)
        {
            Array.Clear(block, position, BlockSize - position);
            Transform(state, block);
        }

        var result = new byte[20];
        for (var i = 0; i < 5; i++)
        {
            BinaryPrimitives.WriteUInt32LittleEndian(result.AsSpan(i * 4), state[i]);
        }

        return result;
    }

    private static void Transform(uint[] state, byte[] block)
    {
        var w = new uint[80];
        for (var i = 0; i < 16; i++)
        {
            w[i] = BinaryPrimitives.ReadUInt32LittleEndian(block.AsSpan(i * 4));
        }

        for (var i = 16; i < 80; i++)
        {
            var value = w[i - 16] ^ w[i - 8] ^ w[i - 14] ^ w[i - 3];
            w[i] = 1u << (int)(value & 31);
        }

        var a = state[0];
        var b = state[1];
        var c = state[2];
        var d = state[3];
        var e = state[4];

        for (var i = 0; i < 80; i++)
        {
            var f = i switch
            {
                < 20 => (b & c) | (~b & d),
                < 40 => b ^ c ^ d,
                < 60 => (b & c) | (b & d) | (c & d),
                _ => b ^ c ^ d,
            };

            var temp = RotateLeft(a, 5) + f + e + w[i] + RoundConstants[i / 20];
            e = d;
            d = c;
            c = RotateLeft(b, 30);
            b = a;
            a = temp;
        }

        state[0] += a;
        state[1] += b;
        state[2] += c;
        state[3] += d;
        state[4] += e;
    }

    private static uint RotateLeft(uint value, int shift) => (value << shift) | (value >> (32 - shift));
}
