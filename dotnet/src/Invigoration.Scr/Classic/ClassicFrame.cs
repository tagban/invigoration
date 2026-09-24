using System.Buffers.Binary;
using Invigoration.Sc2.Protobuf;

namespace Invigoration.Scr.Classic;

/// <summary>
/// The RPC header on SC:R's classic connection. Unlike Front's, every field is
/// a varint, the service hash included, and requests carry a routing value.
/// </summary>
public sealed record ClassicHeader(uint Service, uint Method, uint Token, uint? Routing, uint Size, uint Status, bool IsResponse)
{
    /// <summary>The routing value every client request carries.</summary>
    public const uint RequestRouting = 2525111537; // 0x968224F1

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteUInt32(1, Service);
        w.WriteUInt32(2, Method);
        w.WriteUInt32(3, Token);
        w.WriteUInt32(4, Routing);
        w.WriteUInt32(5, Size);
        w.WriteUInt32(6, Status);
        w.WriteUInt32(9, IsResponse ? 1u : 0u);
        return w.ToArray();
    }

    public static ClassicHeader Decode(byte[] data)
    {
        uint service = 0, method = 0, token = 0, size = 0, status = 0;
        uint? routing = null;
        var isResponse = false;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: service = (uint)r.ReadVarint(); break;
                case 2: method = (uint)r.ReadVarint(); break;
                case 3: token = (uint)r.ReadVarint(); break;
                case 4: routing = (uint)r.ReadVarint(); break;
                case 5: size = (uint)r.ReadVarint(); break;
                case 6: status = (uint)r.ReadVarint(); break;
                case 9: isResponse = r.ReadVarint() != 0; break;
                default: r.Skip(type); break;
            }
        }

        return new ClassicHeader(service, method, token, routing, size, status, isResponse);
    }
}

/// <summary>One RPC on the classic connection: its header and body.</summary>
public sealed record ClassicRpc(ClassicHeader Header, byte[] Body);

/// <summary>
/// Framing on SC:R's classic connection, once a message is unscrambled: one or
/// more RPCs back to back, each a big-endian 2-byte header length, the header,
/// then as many body bytes as the header's size says.
/// </summary>
public static class ClassicFrame
{
    public static byte[] Encode(ClassicHeader header, ReadOnlySpan<byte> body)
    {
        var headerBytes = (header with { Size = (uint)body.Length }).Encode();
        if (headerBytes.Length > ushort.MaxValue)
        {
            throw new InvalidOperationException("Classic RPC header is too long.");
        }

        var frame = new byte[2 + headerBytes.Length + body.Length];
        BinaryPrimitives.WriteUInt16BigEndian(frame, (ushort)headerBytes.Length);
        headerBytes.CopyTo(frame, 2);
        body.CopyTo(frame.AsSpan(2 + headerBytes.Length));
        return frame;
    }

    /// <summary>Splits an unscrambled message into its RPCs. A message that ends partway through one is malformed.</summary>
    public static IReadOnlyList<ClassicRpc> DecodeAll(ReadOnlySpan<byte> message)
    {
        var rpcs = new List<ClassicRpc>();
        var offset = 0;
        while (offset < message.Length)
        {
            if (message.Length - offset < 2)
            {
                throw new InvalidOperationException("Classic message ends inside an RPC's header length.");
            }

            var headerLength = BinaryPrimitives.ReadUInt16BigEndian(message[offset..]);
            offset += 2;
            if (message.Length - offset < headerLength)
            {
                throw new InvalidOperationException("Classic message ends inside an RPC header.");
            }

            var header = ClassicHeader.Decode(message.Slice(offset, headerLength).ToArray());
            offset += headerLength;
            if (message.Length - offset < header.Size)
            {
                throw new InvalidOperationException("Classic message ends inside an RPC body.");
            }

            rpcs.Add(new ClassicRpc(header, message.Slice(offset, (int)header.Size).ToArray()));
            offset += (int)header.Size;
        }

        return rpcs;
    }
}
