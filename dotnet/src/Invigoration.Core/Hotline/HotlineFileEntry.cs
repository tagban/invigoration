using System.Buffers.Binary;
using System.Text;

namespace Invigoration.Core.Hotline;

/// <summary>
/// One row of a directory listing — a file or a folder, as the server describes it in a
/// FileNameWithInfo (200) field.
/// </summary>
public sealed record HotlineFileEntry(string Name, string TypeCode, string CreatorCode, uint Size, bool IsFolder)
{
    /// <summary>Folders report their type as "fldr"; everything else is a file, with the classic Mac type/creator codes real Hotline servers still serve.</summary>
    public const string FolderType = "fldr";

    /// <summary>
    /// Unpacks one FileNameWithInfo field:
    ///
    ///   0    4 bytes  type code ("fldr" for a folder, e.g. "TEXT"/"SIT!" otherwise)
    ///   4    4 bytes  creator code
    ///   8    uint32   size — for a folder this is the number of items inside, not bytes
    ///   12   4 bytes  reserved
    ///   16   uint16   name script (unused here)
    ///   18   uint16   name length
    ///   20   bytes    name
    ///
    /// The offsets are spelled out because getting them wrong is silent: reading the length two
    /// bytes late lands on the first two characters of the name, which for a name that starts with
    /// spaces is 0x2020 — 8224, far past the end — so every row fails its bounds check and the
    /// whole listing comes back empty. That was a real bug, confirmed against MacDomain's own
    /// listing (2026-09-16).
    ///
    /// Null for anything too short or inconsistent, so a listing with one odd row still shows the
    /// rest.
    /// </summary>
    public static HotlineFileEntry? TryParse(ReadOnlySpan<byte> data)
    {
        if (data.Length < 20)
        {
            return null;
        }

        var typeCode = Encoding.ASCII.GetString(data[..4]);
        var creatorCode = Encoding.ASCII.GetString(data.Slice(4, 4));
        var size = BinaryPrimitives.ReadUInt32BigEndian(data[8..]);
        var nameLength = BinaryPrimitives.ReadUInt16BigEndian(data[18..]);
        if (20 + nameLength > data.Length)
        {
            return null;
        }

        var isFolder = typeCode == FolderType;
        return new HotlineFileEntry(
            Encoding.UTF8.GetString(data.Slice(20, nameLength)),
            typeCode,
            creatorCode,
            size,
            isFolder);
    }
}
