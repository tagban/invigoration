using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

/// <summary>
/// Connection-slot records that can arrive after sign-in: keepalives, message
/// frames and the game-site catalog. Layouts are the ones written up on
/// docs.bnet.cc's "StarCraft II chat on Sunken".
/// </summary>
public static class ConnectionRecordDecoder
{
    /// <summary>Ping (10) or Pong (12): a presence bit, then byte-aligned 8 timestamp bytes if present.</summary>
    public static byte[]? DecodeKeepalive(BitReader reader) =>
        reader.Read(1) != 0 ? reader.ReadBytes(8, aligned: true) : null;

    /// <summary>Command 3: one 32-bit value, not needed for chat.</summary>
    public static void SkipCommand3(BitReader reader) => reader.Read(32);

    /// <summary>Command 11: an optional pair of 32-bit values, not needed for chat.</summary>
    public static void SkipCommand11(BitReader reader)
    {
        if (reader.Read(1) == 1)
        {
            reader.Read(32);
            reader.Read(32);
        }
    }

    /// <summary>
    /// GameSiteInfo (14): the client's external IPv4 address and port, then a
    /// 7-bit count of sites, each a name (6-bit length) and an optional
    /// address and port. Replaces an earlier fixed 712-bit skip that only fit
    /// the one capture it was measured on.
    /// </summary>
    public static void SkipGameSiteInfo(BitReader reader)
    {
        reader.Skip(48);
        reader.Align();
        var count = (int)reader.Read(7);
        for (var i = 0; i < count; i++)
        {
            reader.ReadBlob(6);
            if (reader.Read(1) != 0)
            {
                reader.Align();
                reader.Skip(48);
            }
        }
    }

    private static readonly int?[] ArrayArmWidths = [32, 32, 32, 72, 64, null, 136, null];

    /// <summary>
    /// MessageFrame (13): a 14-bit data length (at most 14336), the data, a frame
    /// type, then up to 32 typed headers. Only consumed; chat doesn't use it.
    /// </summary>
    public static void SkipMessageFrame(BitReader reader)
    {
        var size = reader.ReadCount(14, 14336, "Message frame data");
        reader.ReadBytes(size, aligned: true);
        reader.Read(8); // frame type
        var headers = reader.ReadCount(6, 32, "Message frame");
        for (var i = 0; i < headers; i++)
        {
            var kind = reader.Read(8);
            switch (kind)
            {
                case 0: // size, encoding
                    reader.Skip(64);
                    break;
                case 1: // name, hash, command, optional node (label, epoch)
                    reader.Skip(32 + 32 + 6);
                    reader.SkipOptional(r => r.Skip(64));
                    break;
                case 2: // an array of fixed-width items
                    reader.Read(7); // type
                    var arm = (int)reader.Read(3);
                    var items = reader.ReadCount(11, 1024, "Message frame array");
                    var width = ArrayArmWidths[arm]
                        ?? throw new InvalidOperationException($"Message frame array has an unknown item kind {arm}.");
                    reader.Skip(items * width);
                    break;
                case 3: // id, reply
                    reader.Skip(33);
                    break;
                case 5: // label, type, epoch
                    reader.Skip(96);
                    break;
                case 6: // result, message
                    reader.Read(16);
                    reader.ReadBlob(14);
                    break;
                case 7: // command
                    reader.Read(1);
                    break;
                case 8: // seconds
                    reader.Read(32);
                    break;
                case 9: // sequence id, more
                    reader.Skip(17);
                    break;
                default:
                    throw new InvalidOperationException($"Message frame has an unknown header kind {kind}.");
            }
        }
    }
}
