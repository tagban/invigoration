using System.Buffers.Binary;
using System.Text;

namespace Invigoration.Core.Hotline;

/// <summary>
/// The packed path encoding Hotline uses for both a directory (FilePath, 202) and a location in
/// the news tree (NewsPath, 325) — they share a format:
///
///   uint16  component count
///   then per component:
///     uint16  always 0 (two unused/reserved bytes in every real client's encoder)
///     uint8   length of the name
///     bytes   the name
///
/// The root is the empty path, which real clients send by omitting the field entirely rather than
/// writing a zero count — see <see cref="ToFieldOrNull"/>.
/// </summary>
public static class HotlineNewsPath
{
    /// <summary>Packs path components into the wire form. An empty path packs to a bare zero count; callers usually want <see cref="ToFieldOrNull"/>.</summary>
    public static byte[] Encode(IReadOnlyList<string> components)
    {
        var names = components.Select(Encoding.UTF8.GetBytes).ToArray();
        var size = 2 + names.Sum(n => 3 + n.Length);
        var buffer = new byte[size];
        BinaryPrimitives.WriteUInt16BigEndian(buffer, (ushort)names.Length);

        var offset = 2;
        foreach (var name in names)
        {
            // Two reserved bytes, then a single-byte length — a name longer than 255 bytes can't be
            // represented at all, and no real server has one.
            buffer[offset] = 0;
            buffer[offset + 1] = 0;
            buffer[offset + 2] = (byte)Math.Min(name.Length, byte.MaxValue);
            name.AsSpan(0, buffer[offset + 2]).CopyTo(buffer.AsSpan(offset + 3));
            offset += 3 + buffer[offset + 2];
        }

        return buffer;
    }

    /// <summary>The path as a field to add to a request, or null for the root — which is sent by leaving the field out.</summary>
    public static HotlineField? ToFieldOrNull(HotlineFieldType type, IReadOnlyList<string> components) =>
        components.Count == 0 ? null : new HotlineField(type, Encode(components));

    /// <summary>Unpacks a path. Returns an empty list for anything malformed — a path is never worth throwing over.</summary>
    public static IReadOnlyList<string> Decode(ReadOnlySpan<byte> data)
    {
        if (data.Length < 2)
        {
            return [];
        }

        var count = BinaryPrimitives.ReadUInt16BigEndian(data);
        var components = new List<string>(count);
        var offset = 2;
        for (var i = 0; i < count; i++)
        {
            if (offset + 3 > data.Length)
            {
                break;
            }

            var length = data[offset + 2];
            offset += 3;
            if (offset + length > data.Length)
            {
                break;
            }

            components.Add(Encoding.UTF8.GetString(data.Slice(offset, length)));
            offset += length;
        }

        return components;
    }
}
