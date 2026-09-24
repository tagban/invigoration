using System.Buffers.Binary;
using System.Numerics;

namespace Invigoration.Scr.Classic;

/// <summary>
/// The scrambling applied to every message on SC:R's classic WebSocket, keyed
/// by a 32-bit seed agreed during sign-in. Each 4-byte block's key is derived
/// from the scrambled block before it, and every message starts again from the
/// seed. See docs.bnet.cc's "StarCraft: Remastered chat"; the scheme comes from
/// ncarrillo/sc1-research (MIT).
/// </summary>
public static class ClassicEnvelope
{
    /// <summary>Scrambles an outgoing message.</summary>
    public static byte[] Scramble(ReadOnlySpan<byte> plain, uint seed) => Apply(plain, seed, encoding: true);

    /// <summary>Unscrambles an incoming message.</summary>
    public static byte[] Unscramble(ReadOnlySpan<byte> wire, uint seed) => Apply(wire, seed, encoding: false);

    private static byte[] Apply(ReadOnlySpan<byte> input, uint seed, bool encoding)
    {
        var output = input.ToArray();
        var key = seed;
        for (var lane = 0; lane < Math.Min(4, input.Length); lane++)
        {
            key = BitOperations.RotateLeft(key, 1);
            output[lane] ^= (byte)(key >> (lane * 8));
        }

        var offset = 4;
        while (offset < input.Length)
        {
            // Always the scrambled bytes: the input when unscrambling, what we've
            // produced so far when scrambling.
            var scrambled = encoding ? output.AsSpan(offset - 4, 4) : input.Slice(offset - 4, 4);
            var previous = BinaryPrimitives.ReadUInt32LittleEndian(scrambled);
            var rotate = ~(((uint)offset & 31) ^ previous) & 31;
            key = BitOperations.RotateLeft(previous, (int)((32 - rotate) & 31));
            for (var lane = 0; lane < 4 && offset < input.Length; lane++, offset++)
            {
                key = BitOperations.RotateLeft(key, 1);
                output[offset] ^= (byte)(key >> (lane * 8));
            }
        }

        return output;
    }
}
