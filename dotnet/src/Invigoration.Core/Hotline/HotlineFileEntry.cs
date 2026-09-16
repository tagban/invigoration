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
    ///   4 bytes  type code ("fldr" for a folder, e.g. "TEXT"/"SIT!" otherwise)
    ///   4 bytes  creator code
    ///   uint32   size — for a folder this is the number of items inside, not bytes
    ///   4 bytes  reserved
    ///   uint16   name script (unused here)
    ///   uint16   name length
    ///   bytes    name
    ///
    /// Null for anything too short or inconsistent, so a listing with one odd row still shows the
    /// rest.
    /// </summary>
    public static HotlineFileEntry? TryParse(ReadOnlySpan<byte> data)
    {
        if (data.Length < 22)
        {
            return null;
        }

        var typeCode = Encoding.ASCII.GetString(data[..4]);
        var creatorCode = Encoding.ASCII.GetString(data.Slice(4, 4));
        var size = BinaryPrimitives.ReadUInt32BigEndian(data[8..]);
        var nameLength = BinaryPrimitives.ReadUInt16BigEndian(data[20..]);
        if (22 + nameLength > data.Length)
        {
            return null;
        }

        var isFolder = typeCode == FolderType;
        return new HotlineFileEntry(
            Encoding.UTF8.GetString(data.Slice(22, nameLength)),
            typeCode,
            creatorCode,
            size,
            isFolder);
    }
}
