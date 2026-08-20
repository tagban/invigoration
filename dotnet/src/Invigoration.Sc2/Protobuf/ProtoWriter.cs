using System.Buffers.Binary;
using System.Text;

namespace Invigoration.Sc2.Protobuf;

/// <summary>
/// Minimal proto2 wire encoder for the fixed, known set of Battle.net Front
/// messages this client speaks. Not a general-purpose protobuf library:
/// there is no support for zigzag (sint) fields or packed repeated scalars,
/// since none of the messages this client sends or reads use them.
/// </summary>
public sealed class ProtoWriter
{
    private readonly MemoryStream _stream = new();

    public byte[] ToArray() => _stream.ToArray();

    public void WriteVarint(ulong value)
    {
        while (value >= 0x80)
        {
            _stream.WriteByte((byte)(value | 0x80));
            value >>= 7;
        }

        _stream.WriteByte((byte)value);
    }

    private void WriteTag(int field, WireType type) => WriteVarint((ulong)((field << 3) | (int)type));

    public void WriteUInt64(int field, ulong? value)
    {
        if (value is null)
        {
            return;
        }

        WriteTag(field, WireType.Varint);
        WriteVarint(value.Value);
    }

    public void WriteUInt32(int field, uint? value) => WriteUInt64(field, value);

    public void WriteInt64(int field, long? value)
    {
        if (value is null)
        {
            return;
        }

        WriteTag(field, WireType.Varint);
        WriteVarint(unchecked((ulong)value.Value));
    }

    public void WriteInt32(int field, int? value) => WriteInt64(field, value);

    public void WriteBool(int field, bool? value)
    {
        if (value is null)
        {
            return;
        }

        WriteTag(field, WireType.Varint);
        WriteVarint(value.Value ? 1u : 0u);
    }

    public void WriteFixed32(int field, uint? value)
    {
        if (value is null)
        {
            return;
        }

        WriteTag(field, WireType.Fixed32);
        Span<byte> bytes = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32LittleEndian(bytes, value.Value);
        _stream.Write(bytes);
    }

    public void WriteFixed64(int field, ulong? value)
    {
        if (value is null)
        {
            return;
        }

        WriteTag(field, WireType.Fixed64);
        Span<byte> bytes = stackalloc byte[8];
        BinaryPrimitives.WriteUInt64LittleEndian(bytes, value.Value);
        _stream.Write(bytes);
    }

    public void WriteDouble(int field, double? value)
    {
        if (value is null)
        {
            return;
        }

        WriteTag(field, WireType.Fixed64);
        Span<byte> bytes = stackalloc byte[8];
        BinaryPrimitives.WriteUInt64LittleEndian(bytes, BitConverter.DoubleToUInt64Bits(value.Value));
        _stream.Write(bytes);
    }

    public void WriteString(int field, string? value)
    {
        if (value is null)
        {
            return;
        }

        WriteBytesField(field, Encoding.UTF8.GetBytes(value));
    }

    public void WriteBytesField(int field, byte[]? value)
    {
        if (value is null)
        {
            return;
        }

        WriteTag(field, WireType.LengthDelimited);
        WriteVarint((ulong)value.Length);
        _stream.Write(value);
    }
}
