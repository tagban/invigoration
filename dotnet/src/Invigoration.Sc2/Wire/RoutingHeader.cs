namespace Invigoration.Sc2.Wire;

/// <summary>
/// The 7-or-11-bit prefix on every native ("Sunken") record: a 6-bit command
/// id, a 1-bit "does a service slot follow" flag, and an optional 4-bit
/// service slot. Ported from core/src/bsn/bits.rs's encode_routing_header /
/// strip_routing_header.
/// </summary>
public readonly record struct RoutingHeader(byte CommandId, byte? ServiceSlot, int BitCount)
{
    public static RoutingHeader Encode(BitWriter writer, byte commandId, byte? serviceSlot)
    {
        if (commandId > 0x3f)
        {
            throw new ArgumentOutOfRangeException(nameof(commandId), "Command id must fit in 6 bits.");
        }

        if (serviceSlot is { } slot && slot > 0x0f)
        {
            throw new ArgumentOutOfRangeException(nameof(serviceSlot), "Service slot must fit in 4 bits.");
        }

        writer.Write(commandId, 6);
        writer.Write((ulong)(serviceSlot is null ? 0 : 1), 1);
        if (serviceSlot is { } value)
        {
            writer.Write(value, 4);
        }

        return new RoutingHeader(commandId, serviceSlot, serviceSlot is null ? 7 : 11);
    }

    public static RoutingHeader Decode(BitReader reader)
    {
        var startPosition = reader.Position;
        var commandId = (byte)reader.Read(6);
        var hasSlot = reader.Read(1) != 0;
        byte? slot = null;
        if (hasSlot)
        {
            slot = (byte)reader.Read(4);
        }

        return new RoutingHeader(commandId, slot, reader.Position - startPosition);
    }
}
