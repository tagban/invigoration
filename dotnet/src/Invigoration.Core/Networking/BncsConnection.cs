using Invigoration.Core.Protocol;

namespace Invigoration.Core.Networking;

/// <summary>BNCS framing: FF, id, WORD length (LE, includes this 4-byte header), payload.</summary>
public sealed class BncsConnection : FramedTcpClient
{
    protected override int? TryGetFrameLength(IReadOnlyList<byte> buffer)
    {
        if (buffer.Count < 4)
        {
            return null;
        }

        return buffer[2] | (buffer[3] << 8);
    }

    public Task SendAsync(PacketWriter writer, BncsPacketId id, CancellationToken cancellationToken = default)
        => SendAsync(writer.ToBncsPacket(id), cancellationToken);

    public static byte GetPacketId(byte[] frame) => frame[1];

    public static PacketReader GetPayloadReader(byte[] frame) => new(frame, offset: 4);
}
