using Invigoration.Core.Protocol;

namespace Invigoration.Core.Networking;

/// <summary>
/// BNLS framing: WORD length (LE, includes this 3-byte header), id, payload.
/// Unlike the VB6 original (which parsed whatever a single DataArrival chunk
/// happened to contain, with no length-based buffering), this properly
/// reassembles frames split or coalesced across TCP reads.
/// </summary>
public sealed class BnlsConnection : FramedTcpClient
{
    protected override int? TryGetFrameLength(IReadOnlyList<byte> buffer)
    {
        if (buffer.Count < 3)
        {
            return null;
        }

        return buffer[0] | (buffer[1] << 8);
    }

    public Task SendAsync(PacketWriter writer, BnlsPacketId id, CancellationToken cancellationToken = default)
        => SendAsync(writer.ToBnlsPacket(id), cancellationToken);

    public static byte GetPacketId(byte[] frame) => frame[2];

    public static PacketReader GetPayloadReader(byte[] frame) => new(frame, offset: 3);
}
