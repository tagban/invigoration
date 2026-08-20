using System.Buffers.Binary;

namespace Invigoration.Sc2.Front;

/// <summary>
/// Wire framing for the Front WebSocket transport: `u16 BE header_length ++
/// Header ++ body`, per core/src/wire/protobuf.rs. Verified against the
/// crate's own vector for {service_id:0, method_id:1, token:0, size:2,
/// service_hash:0x65446991} + ConnectRequest{use_bindless_rpc:true}:
/// "000d08001001180028025d916944651801".
/// </summary>
public static class FrontFrame
{
    public static byte[] Encode(Header header, byte[] body)
    {
        var headerBytes = header.Encode();
        var frame = new byte[2 + headerBytes.Length + body.Length];
        BinaryPrimitives.WriteUInt16BigEndian(frame, (ushort)headerBytes.Length);
        headerBytes.CopyTo(frame, 2);
        body.CopyTo(frame, 2 + headerBytes.Length);
        return frame;
    }

    public static (Header Header, byte[] Body) Decode(byte[] frame)
    {
        var headerLength = BinaryPrimitives.ReadUInt16BigEndian(frame);
        var header = Header.Decode(frame.AsSpan(2, headerLength).ToArray());
        var body = frame.AsSpan(2 + headerLength).ToArray();
        return (header, body);
    }
}
