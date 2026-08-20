using System.Buffers.Binary;
using System.Text;

namespace Invigoration.Sc2.Protobuf;

/// <summary>Counterpart to <see cref="ProtoWriter"/>. See its remarks for scope.</summary>
public sealed class ProtoReader
{
    private readonly byte[] _data;
    private int _pos;

    public ProtoReader(byte[] data) => _data = data;

    public bool HasMore => _pos < _data.Length;

    public (int Field, WireType Type) ReadTag()
    {
        var tag = ReadVarint();
        return ((int)(tag >> 3), (WireType)(tag & 0x7));
    }

    public ulong ReadVarint()
    {
        ulong result = 0;
        var shift = 0;
        while (true)
        {
            var b = _data[_pos++];
            result |= (ulong)(b & 0x7f) << shift;
            if ((b & 0x80) == 0)
            {
                break;
            }

            shift += 7;
        }

        return result;
    }

    public uint ReadFixed32()
    {
        var value = BinaryPrimitives.ReadUInt32LittleEndian(_data.AsSpan(_pos, 4));
        _pos += 4;
        return value;
    }

    public ulong ReadFixed64()
    {
        var value = BinaryPrimitives.ReadUInt64LittleEndian(_data.AsSpan(_pos, 8));
        _pos += 8;
        return value;
    }

    public double ReadDouble() => BitConverter.UInt64BitsToDouble(ReadFixed64());

    public byte[] ReadLengthDelimited()
    {
        var len = (int)ReadVarint();
        var result = _data.AsSpan(_pos, len).ToArray();
        _pos += len;
        return result;
    }

    public string ReadString() => Encoding.UTF8.GetString(ReadLengthDelimited());

    public void Skip(WireType type)
    {
        switch (type)
        {
            case WireType.Varint:
                ReadVarint();
                break;
            case WireType.Fixed64:
                _pos += 8;
                break;
            case WireType.LengthDelimited:
                ReadLengthDelimited();
                break;
            case WireType.Fixed32:
                _pos += 4;
                break;
            default:
                throw new InvalidOperationException($"Unsupported wire type {type}");
        }
    }
}
