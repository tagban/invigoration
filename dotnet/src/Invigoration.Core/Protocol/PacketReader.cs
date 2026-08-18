using System.Text;

namespace Invigoration.Core.Protocol;

/// <summary>
/// Sequential little-endian reader over a packet payload. Replaces the VB6
/// Buffer.cls, which did the same thing via CopyMemory over a byte-per-char string.
/// </summary>
public sealed class PacketReader
{
    private readonly byte[] _buffer;
    private int _position;

    public PacketReader(byte[] buffer, int offset = 0)
    {
        _buffer = buffer;
        _position = offset;
    }

    public int Position => _position;

    public int Remaining => _buffer.Length - _position;

    public void Skip(int count) => _position += count;

    public byte ReadByte()
    {
        var value = _buffer[_position];
        _position += 1;
        return value;
    }

    public ushort ReadWord()
    {
        var value = (ushort)(_buffer[_position] | (_buffer[_position + 1] << 8));
        _position += 2;
        return value;
    }

    public uint ReadDword()
    {
        var value = (uint)(_buffer[_position]
            | (_buffer[_position + 1] << 8)
            | (_buffer[_position + 2] << 16)
            | (_buffer[_position + 3] << 24));
        _position += 4;
        return value;
    }

    /// <summary>
    /// BNLS booleans are sent as a full DWORD (0 = false, nonzero = true),
    /// matching Buffer.cls.GetBoolean.
    /// </summary>
    public bool ReadBoolean() => ReadDword() != 0;

    public byte[] ReadRaw(int length)
    {
        var result = new byte[length];
        Array.Copy(_buffer, _position, result, 0, length);
        _position += length;
        return result;
    }

    /// <summary>
    /// Reads a Win32 FILETIME (two little-endian DWORDs: low, then high).
    /// </summary>
    public FileTimeValue ReadFileTime()
    {
        var low = ReadDword();
        var high = ReadDword();
        return new FileTimeValue(low, high);
    }

    /// <summary>Reads a null-terminated ASCII string, consuming the terminator.</summary>
    public string ReadNTString()
    {
        var start = _position;
        while (_position < _buffer.Length && _buffer[_position] != 0)
        {
            _position++;
        }

        var value = Encoding.Latin1.GetString(_buffer, start, _position - start);
        if (_position < _buffer.Length)
        {
            _position++; // consume the null terminator
        }

        return value;
    }
}

/// <summary>A raw Win32 FILETIME (low/high DWORD pair), kept unparsed since callers echo it back verbatim.</summary>
public readonly record struct FileTimeValue(uint Low, uint High);
