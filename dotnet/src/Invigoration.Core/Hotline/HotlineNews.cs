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
    ///   uint16  count — articles in a category, items in a bundle
    ///   (categories only) 16-byte GUID + uint32 add-SN + uint32 delete-SN, unused here
    ///   uint8   name length
    ///   bytes   name
    ///
    /// A bundle stops after the count, so the two are sized separately — and the GUID is 16 bytes,
    /// not 8. Confirmed byte-for-byte against MacDomain (2026-09-16), whose reply is:
    ///
    ///   0002 0005 05 "Files"                          — a bundle, 10 bytes
    ///   0003 005b [24 zero bytes] 09 "Guestbook"      — a category, 38 bytes
    ///
    /// Sizing the category's fixed part at 20 instead of 28 reads the name length out of the
    /// middle of the GUID, which lands on a zero, and every entry after it is lost — which is why
    /// this server's news looked empty.
    ///
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

            // Bundle: type + count, then the name. Category: the same plus a 16-byte GUID and two serials.
            var fixedPart = isBundle ? 4 : 4 + 16 + 4 + 4;
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
                count));
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
    ///   then per article, a 22-byte header:
    ///     +0   uint32  article ID
    ///     +4   8 bytes date (year, ms, seconds-into-year)
    ///     +12  uint32  parent ID — non-zero makes this a reply
    ///     +16  uint16  flags
    ///     +18  uint16  reserved; every server seen sends zero
    ///     +20  uint16  flavour count
    ///   then:
    ///     uint8   title length, title
    ///     uint8   poster length, poster
    ///     then per flavour: uint8 length + name, uint16 size
    ///
    /// That header is 22 bytes, not the 20 the docs' field list implies. Confirmed against
    /// MacDomain's Guestbook (2026-09-16), where article 1 reads
    /// <c>00000001 | 07e6 0000 01dc2a48 | 00000000 | 0000 0000 0001 | 0a "Greetings!" 0a "MacDude888"</c> —
    /// the flavour count is plainly the third uint16 after the parent, not the first. Reading it
    /// two bytes early takes the count as 0 and then the title length from a zero byte, so the
    /// first article comes out blank and the walk lands mid-way through the second: 91 articles
    /// parsed as one empty row.
    ///
    /// The flavour list is the other thing that bites: it's variable-length and sits at the END of
    /// each article, so skipping it wrong walks into the middle of the next one. Parsing stops at
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
            if (offset + 22 > data.Length)
            {
                break;
            }

            var id = BinaryPrimitives.ReadUInt32BigEndian(data[offset..]);
            var posted = ParseDate(data.Slice(offset + 4, 8));
            var parentId = BinaryPrimitives.ReadUInt32BigEndian(data[(offset + 12)..]);
            var flavorCount = BinaryPrimitives.ReadUInt16BigEndian(data[(offset + 20)..]);
            offset += 22;

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
