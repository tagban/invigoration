using Invigoration.Core.Protocol;

namespace Invigoration.Core.Networking;

/// <summary>
/// D2 realm server connection (uses the same 3-byte-header framing as BNLS).
/// The original VB6 bot connects here on SID_LOGONREALMEX (0x3E) but never
/// registered a DataArrival handler, so no realm packet was ever parsed —
/// that gap is preserved here rather than inventing new protocol behavior.
/// </summary>
public sealed class RealmConnection : FramedTcpClient
{
    protected override int? TryGetFrameLength(IReadOnlyList<byte> buffer)
    {
        if (buffer.Count < 3)
        {
            return null;
        }

        return buffer[0] | (buffer[1] << 8);
    }

    public Task SendAsync(PacketWriter writer, byte id, CancellationToken cancellationToken = default)
        => SendAsync(writer.ToRealmPacket(id), cancellationToken);

    public static byte GetPacketId(byte[] frame) => frame[2];

    public static PacketReader GetPayloadReader(byte[] frame) => new(frame, offset: 3);
}
