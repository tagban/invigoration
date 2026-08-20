using System.Text;

namespace Invigoration.Sc2.Wire;

/// <summary>
/// Battle.net's "FourCC" packing: big-endian bytes folded into a uint32.
/// Ported from core/src/native/protocol.rs's fourcc() — note it left-pads
/// (not right-pads) strings shorter than 4 characters, since the fold starts
/// from an accumulator of 0 and shifts left each byte in: fourcc("S2") ==
/// 0x00005332, not 0x53320000.
/// </summary>
public static class FourCc
{
    public static uint Encode(string value)
    {
        var bytes = Encoding.ASCII.GetBytes(value);
        uint result = 0;
        foreach (var b in bytes)
        {
            result = (result << 8) | b;
        }

        return result;
    }
}
