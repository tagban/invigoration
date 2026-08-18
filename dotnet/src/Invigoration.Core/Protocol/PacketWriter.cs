using System.Text;

namespace Invigoration.Core.Protocol;

/// <summary>
/// Builds a single outbound packet payload, then frames it for BNCS, BNLS, or
/// the D2 realm server. Replaces the VB6 packetbuffer.cls, but as a one-shot
/// builder per packet rather than a shared mutable singleton (PBuffer) that
/// had to be cleared after every send.
/// </summary>
public sealed class PacketWriter
{
    private readonly List<byte> _buffer = [];

    public PacketWriter WriteByte(byte value)
    {
        _buffer.Add(value);
        return this;
    }

    public PacketWriter WriteWord(ushort value)
    {
        _buffer.Add((byte)(value & 0xFF));
        _buffer.Add((byte)((value >> 8) & 0xFF));
        return this;
    }

    public PacketWriter WriteDword(uint value)
    {
        _buffer.Add((byte)(value & 0xFF));
        _buffer.Add((byte)((value >> 8) & 0xFF));
        _buffer.Add((byte)((value >> 16) & 0xFF));
        _buffer.Add((byte)((value >> 24) & 0xFF));
        return this;
    }

    public PacketWriter WriteBytes(ReadOnlySpan<byte> data)
    {
        _buffer.AddRange(data.ToArray());
        return this;
    }

    /// <summary>Writes raw ASCII text with no terminator (VB6 InsertNonNTString).</summary>
    public PacketWriter WriteAscii(string text)
    {
        _buffer.AddRange(Encoding.Latin1.GetBytes(text));
        return this;
    }

    /// <summary>Writes ASCII text followed by a null terminator (VB6 InsertNTString).</summary>
    public PacketWriter WriteNTString(string text)
    {
        WriteAscii(text);
        _buffer.Add(0);
        return this;
    }

    /// <summary>Frames the payload as a BNCS packet: FF, id, WORD length (LE, includes this 4-byte header), payload.</summary>
    public byte[] ToBncsPacket(BncsPacketId id)
    {
        var length = (ushort)(_buffer.Count + 4);
        var result = new byte[length];
        result[0] = 0xFF;
        result[1] = (byte)id;
        result[2] = (byte)(length & 0xFF);
        result[3] = (byte)((length >> 8) & 0xFF);
        _buffer.CopyTo(result, 4);
        return result;
    }

    /// <summary>Frames the payload as a BNLS packet: WORD length (LE, includes this 3-byte header), id, payload.</summary>
    public byte[] ToBnlsPacket(BnlsPacketId id) => ToThreeByteHeaderPacket((byte)id);

    /// <summary>Frames the payload as a D2 realm packet: same 3-byte-header layout as BNLS.</summary>
    public byte[] ToRealmPacket(byte id) => ToThreeByteHeaderPacket(id);

    private byte[] ToThreeByteHeaderPacket(byte id)
    {
        var length = (ushort)(_buffer.Count + 3);
        var result = new byte[length];
        result[0] = (byte)(length & 0xFF);
        result[1] = (byte)((length >> 8) & 0xFF);
        result[2] = id;
        _buffer.CopyTo(result, 3);
        return result;
    }
}
