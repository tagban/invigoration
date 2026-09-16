using System.Buffers.Binary;
using System.Reflection;
using Invigoration.Core.Hotline;

namespace Invigoration.Core.Tests;

/// <summary>
/// Bytes captured from a real server (MacDomain, as a guest, 2026-09-16) rather than bytes this
/// code made up. Two bugs got through a full set of self-built fixtures because the fixtures
/// encoded the same wrong layout the parser decoded — these are the antidote.
/// </summary>
public class HotlineRealServerFixtureTests
{
    /// <summary>A real FileNameWithInfo row: a text file whose name begins with spaces, which is what exposed the offset bug.</summary>
    private const string RealTextFileEntry =
        "54455854747478740000000f000000000000002c2020202020736565204e45575320666f7220496e666f2061626f75742046696c65732020202020202e747874";

    /// <summary>A real folder row.</summary>
    private const string RealFolderEntry =
        "666c647200000000000000020000000000000015202020666f7220796f757220646f776e6c6f616473";

    /// <summary>MacDomain's actual access bitmap for a guest account.</summary>
    private const string RealGuestAccessBitmap = "60700c2003800000";

    [Fact]
    public void AFileRowFromARealServer_ParsesWithItsNameIntact()
    {
        var entry = HotlineFileEntry.TryParse(Convert.FromHexString(RealTextFileEntry));

        Assert.NotNull(entry);
        Assert.Equal("TEXT", entry.TypeCode);
        Assert.False(entry.IsFolder);
        Assert.Equal(15u, entry.Size);

        // Leading spaces are the server's own formatting — kept, not trimmed, and crucially not
        // mistaken for the name's length.
        Assert.Equal("     see NEWS for Info about Files      .txt", entry.Name);
    }

    [Fact]
    public void AFolderRowFromARealServer_Parses()
    {
        var entry = HotlineFileEntry.TryParse(Convert.FromHexString(RealFolderEntry));

        Assert.NotNull(entry);
        Assert.True(entry.IsFolder);
        Assert.Equal(2u, entry.Size); // items inside, not bytes
        Assert.Equal("   for your downloads", entry.Name);
    }

    /// <summary>
    /// The exact shape of the bug: a name starting with two spaces reads as a length of 0x2020
    /// when taken from the wrong offset, which fails the bounds check and drops the row — so an
    /// entire listing silently became "no files".
    /// </summary>
    [Fact]
    public void EveryRowOfARealListing_Survives()
    {
        foreach (var hex in new[] { RealTextFileEntry, RealFolderEntry })
        {
            Assert.NotNull(HotlineFileEntry.TryParse(Convert.FromHexString(hex)));
        }
    }

    private static HotlineTransactionClient ClientWithAccessBits(string hex)
    {
        var client = new HotlineTransactionClient();
        typeof(HotlineTransactionClient)
            .GetProperty(nameof(HotlineTransactionClient.OwnAccessBits))!
            .SetValue(client, BinaryPrimitives.ReadUInt64BigEndian(Convert.FromHexString(hex)));
        return client;
    }

    /// <summary>
    /// A guest on MacDomain can download, upload, read and post news. Reading the bitmap the other
    /// way round denied all four — which is what produced "this account isn't allowed to browse
    /// files" for an account that plainly was.
    /// </summary>
    [Fact]
    public async Task ARealGuestAccount_HasThePrivilegesTheServerActuallyGrantedIt()
    {
        await using var client = ClientWithAccessBits(RealGuestAccessBitmap);

        Assert.True(client.CanDownloadFiles, "a guest on MacDomain can browse and download");
        Assert.True(client.CanUploadFiles);
        Assert.True(client.CanReadNews);
        Assert.True(client.CanPostNews);
    }

    /// <summary>The same bitmap must NOT grant what it doesn't — a check that the bit order isn't just flipped into being permissive.</summary>
    [Fact]
    public async Task ARealGuestAccount_DoesNotHaveAdministrativePrivileges()
    {
        await using var client = ClientWithAccessBits(RealGuestAccessBitmap);

        Assert.False(client.HasOwnAccess(HotlineAccessBits.CreateUser));
        Assert.False(client.HasOwnAccess(HotlineAccessBits.DeleteUser));
        Assert.False(client.HasOwnAccess(HotlineAccessBits.DisconnectUser));
        Assert.False(client.HasOwnAccess(HotlineAccessBits.CannotBeDisconnected));
        Assert.False(client.HasOwnAccess(HotlineAccessBits.DeleteFile));
    }

    /// <summary>
    /// MacDomain's real category listing: a bundle then a category. The category's fixed part is 28
    /// bytes — a 16-byte GUID, not 8 — and sizing it at 20 reads the name length out of the middle
    /// of the GUID (a zero byte), losing that entry and everything after it. That is why this
    /// server's news looked empty.
    /// </summary>
    [Fact]
    public void ARealNewsCategoryListing_ParsesBothEntries()
    {
        var data = Convert.FromHexString(
            "000200050546696c6573" +
            "0003005b000000000000000000000000000000000000000000000000094775657374626f6f6b");

        var categories = HotlineNewsCategory.ParseList(data);

        Assert.Equal(2, categories.Count);

        Assert.Equal("Files", categories[0].Name);
        Assert.True(categories[0].IsBundle);
        Assert.Equal(5, categories[0].ArticleCount);

        Assert.Equal("Guestbook", categories[1].Name);
        Assert.False(categories[1].IsBundle);
        Assert.Equal(91, categories[1].ArticleCount);
    }

    /// <summary>
    /// MacDomain's real Guestbook listing. The article header is 22 bytes; reading the flavour
    /// count two bytes early (the 20 the docs imply) takes it as zero and then reads the title
    /// length from a zero byte, so 91 articles came out as a single blank row.
    /// </summary>
    [Fact]
    public void ARealArticleListing_ParsesTitlesPostersAndReplies()
    {
        var data = Convert.FromHexString(
            "00000000" + "0000005b" + "0000" +                                     // first id, count 91, no name
            "00000001" + "07e6000001dc2a48" + "00000000" + "000000000001" +        // article 1 header
            "0a" + "4772656574696e677321" + "0a" + "4d6163447564653838380a746578742f706c61696e0025" +
            "00000003" + "07e7000000245b1f" + "00000002" + "000000000001" +        // article 3, a reply to 2
            "17" + "52653a2048692066726f6d2050697462756c6c2050726f" +
            "07" + "4b6e657a7a656e" + "0a746578742f706c61696e0010");

        var articles = HotlineNewsArticle.ParseList(data);

        Assert.Equal(2, articles.Count);

        Assert.Equal(1u, articles[0].Id);
        Assert.Equal("Greetings!", articles[0].Title);
        Assert.Equal("MacDude888", articles[0].Poster);
        Assert.Equal(0u, articles[0].ParentId);
        Assert.Equal("text/plain", articles[0].Flavor);
        Assert.Equal(2022, articles[0].Posted?.Year);

        // The reply, which only parses if the first article's flavour list was walked correctly.
        Assert.Equal(3u, articles[1].Id);
        Assert.Equal("Re: Hi from Pitbull Pro", articles[1].Title);
        Assert.Equal("Knezzen", articles[1].Poster);
        Assert.Equal(2u, articles[1].ParentId);
    }

    /// <summary>Bit 0 is the top bit of the first byte — the single fact the protocol docs never state.</summary>
    [Theory]
    [InlineData("8000000000000000", 0)]
    [InlineData("4000000000000000", 1)]
    [InlineData("2000000000000000", 2)]
    [InlineData("0000000000000001", 63)]
    public async Task BitNumbering_RunsFromTheTopBitOfTheFirstByte(string hex, int expectedBit)
    {
        await using var client = ClientWithAccessBits(hex);

        for (var bit = 0; bit < 64; bit++)
        {
            Assert.Equal(bit == expectedBit, client.HasOwnAccess(bit));
        }
    }
}
