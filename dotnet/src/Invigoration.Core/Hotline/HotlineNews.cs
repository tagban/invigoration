using System.Buffers.Binary;
using System.Text;

namespace Invigoration.Core.Hotline;

/// <summary>One entry in the news tree: either a bundle (a folder of categories) or a category (a folder of articles).</summary>
public sealed record HotlineNewsCategory(string Name, bool IsBundle, ushort ArticleCount)
{
    /// <summary>
    /// Unpacks the reply to GetNewsCategoryNameList (370). Each entry is:
    ///
    ///   uint16  type — 2 = bundle, 3 = category
    ///   uint16  article count (categories only; a bundle's is meaningless)
    ///   bytes   8-byte GUID + uint32 add-SN + uint32 delete-SN, all unused here
    ///   uint8   name length
    ///   bytes   name
    ///
    /// A bundle carries no counts or serial numbers, so its fixed part is shorter — getting that
    /// wrong desynchronizes the whole list, which is why the two are sized separately below.
    /// Stops at the first entry that doesn't fit rather than throwing: a listing this client can't
    /// fully parse is still worth showing as far as it got.
    /// </summary>
    public static IReadOnlyList<HotlineNewsCategory> ParseList(ReadOnlySpan<byte> data)
    {
        var categories = new List<HotlineNewsCategory>();
        var offset = 0;

        while (offset + 4 <= data.Length)
        {
            var type = BinaryPrimitives.ReadUInt16BigEndian(data[offset..]);
            var isBundle = type == 2;
            var count = BinaryPrimitives.ReadUInt16BigEndian(data[(offset + 2)..]);

            // Bundle: type + count, then the name. Category: the same plus GUID and two serials.
            var fixedPart = isBundle ? 4 : 4 + 8 + 4 + 4;
            if (offset + fixedPart + 1 > data.Length)
            {
                break;
            }

            var nameLength = data[offset + fixedPart];
            var nameStart = offset + fixedPart + 1;
            if (nameStart + nameLength > data.Length)
            {
                break;
            }

            categories.Add(new HotlineNewsCategory(
                Encoding.UTF8.GetString(data.Slice(nameStart, nameLength)),
                isBundle,
                isBundle ? (ushort)0 : count));
            offset = nameStart + nameLength;
        }

        return categories;
    }
}

/// <summary>One article's headline row — everything the list shows before you open it.</summary>
public sealed record HotlineNewsArticle(
    uint Id,
    string Title,
    string Poster,
    DateTimeOffset? Posted,
    uint ParentId,
    string Flavor)
{
    /// <summary>
    /// Unpacks the reply to GetNewsArticleNameList (371):
    ///
    ///   uint32  ID of the first article
    ///   uint32  count
    ///   uint16  name length, then the name
    ///   then per article:
    ///     uint32  article ID
    ///     8 bytes date, uint32 parent ID, uint16 flags
    ///     uint16  flavour count
    ///     uint8   title length, title
    ///     uint8   poster length, poster
    ///     then per flavour: uint8 length + name, uint16 size
    ///
    /// The flavour list is the part that bites: it's variable-length and sits at the END of each
    /// article, so skipping it wrong walks into the middle of the next article. Parsing stops at
    /// the first entry that doesn't fit rather than throwing.
    /// </summary>
    public static IReadOnlyList<HotlineNewsArticle> ParseList(ReadOnlySpan<byte> data)
    {
        var articles = new List<HotlineNewsArticle>();
        if (data.Length < 10)
        {
            return articles;
        }

        var count = BinaryPrimitives.ReadUInt32BigEndian(data[4..]);
        var nameLength = BinaryPrimitives.ReadUInt16BigEndian(data[8..]);
        var offset = 10 + nameLength;

        for (var i = 0u; i < count; i++)
        {
            if (offset + 20 > data.Length)
            {
                break;
            }

            var id = BinaryPrimitives.ReadUInt32BigEndian(data[offset..]);
            var posted = ParseDate(data.Slice(offset + 4, 8));
            var parentId = BinaryPrimitives.ReadUInt32BigEndian(data[(offset + 12)..]);
            var flavorCount = BinaryPrimitives.ReadUInt16BigEndian(data[(offset + 18)..]);
            offset += 20;

            if (!TryReadShortString(data, ref offset, out var title) ||
                !TryReadShortString(data, ref offset, out var poster))
            {
                break;
            }

            var flavor = "";
            for (var f = 0; f < flavorCount; f++)
            {
                if (!TryReadShortString(data, ref offset, out var flavorName) || offset + 2 > data.Length)
                {
                    return articles;
                }

                flavor = f == 0 ? flavorName : flavor;
                offset += 2; // the flavour's size, which the body request doesn't need
            }

            articles.Add(new HotlineNewsArticle(id, title, poster, posted, parentId, flavor));
        }

        return articles;
    }

    private static bool TryReadShortString(ReadOnlySpan<byte> data, ref int offset, out string value)
    {
        value = "";
        if (offset >= data.Length)
        {
            return false;
        }

        var length = data[offset];
        if (offset + 1 + length > data.Length)
        {
            return false;
        }

        value = Encoding.UTF8.GetString(data.Slice(offset + 1, length));
        offset += 1 + length;
        return true;
    }

    /// <summary>
    /// Hotline's 8-byte date: uint16 year, uint16 milliseconds (ignored — no real server sets it
    /// meaningfully), uint32 seconds into that year. Null for a date that doesn't make sense,
    /// which is better than showing a nonsense timestamp.
    /// </summary>
    internal static DateTimeOffset? ParseDate(ReadOnlySpan<byte> data)
    {
        if (data.Length < 8)
        {
            return null;
        }

        var year = BinaryPrimitives.ReadUInt16BigEndian(data);
        var seconds = BinaryPrimitives.ReadUInt32BigEndian(data[4..]);
        if (year is < 1900 or > 9999)
        {
            return null;
        }

        try
        {
            return new DateTimeOffset(year, 1, 1, 0, 0, 0, TimeSpan.Zero).AddSeconds(seconds);
        }
        catch (ArgumentOutOfRangeException)
        {
            return null;
        }
    }
}
