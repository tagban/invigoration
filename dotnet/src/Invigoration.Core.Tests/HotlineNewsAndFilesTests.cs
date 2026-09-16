using System.Buffers.Binary;
using System.Net;
using System.Net.Sockets;
using System.Text;
using Invigoration.Core.Hotline;

namespace Invigoration.Core.Tests;

/// <summary>
/// The packed path both file and news requests use. Getting the two reserved bytes or the
/// single-byte length wrong makes a server answer about the wrong folder rather than erroring, so
/// these pin the exact bytes rather than only round-tripping.
/// </summary>
public class HotlineNewsPathTests
{
    [Fact]
    public void Encode_WritesCountThenReservedBytesThenLengthPrefixedNames()
    {
        var encoded = HotlineNewsPath.Encode(["News", "General"]);

        Assert.Equal(2, BinaryPrimitives.ReadUInt16BigEndian(encoded));
        Assert.Equal(0, encoded[2]);
        Assert.Equal(0, encoded[3]);
        Assert.Equal(4, encoded[4]);
        Assert.Equal("News", Encoding.UTF8.GetString(encoded, 5, 4));
        Assert.Equal(7, encoded[11]);
        Assert.Equal("General", Encoding.UTF8.GetString(encoded, 12, 7));
    }

    [Theory]
    [InlineData("News")]
    [InlineData("News", "General", "Deep")]
    public void EncodeThenDecode_RoundTrips(params string[] components)
    {
        Assert.Equal(components, HotlineNewsPath.Decode(HotlineNewsPath.Encode(components)));
    }

    /// <summary>The root is sent by leaving the field out entirely, which is what real clients do.</summary>
    [Fact]
    public void AnEmptyPath_ProducesNoFieldAtAll()
    {
        Assert.Null(HotlineNewsPath.ToFieldOrNull(HotlineFieldType.NewsPath, []));
        Assert.NotNull(HotlineNewsPath.ToFieldOrNull(HotlineFieldType.NewsPath, ["News"]));
    }

    [Theory]
    [InlineData(new byte[] { 0 })]              // too short for a count
    [InlineData(new byte[] { 0, 5, 0, 0, 9 })]  // claims 5 components, has none
    public void Malformed_DecodesToWhateverWasReadable(byte[] data) =>
        Assert.Empty(HotlineNewsPath.Decode(data));
}

public class HotlineNewsParsingTests
{
    private static byte[] Category(string name, bool isBundle, ushort articleCount)
    {
        var nameBytes = Encoding.UTF8.GetBytes(name);
        // A bundle has no GUID or serial numbers — the whole point of the size difference.
        var fixedPart = isBundle ? 4 : 20;
        var entry = new byte[fixedPart + 1 + nameBytes.Length];
        BinaryPrimitives.WriteUInt16BigEndian(entry, isBundle ? (ushort)2 : (ushort)3);
        BinaryPrimitives.WriteUInt16BigEndian(entry.AsSpan(2), articleCount);
        entry[fixedPart] = (byte)nameBytes.Length;
        nameBytes.CopyTo(entry.AsSpan(fixedPart + 1));
        return entry;
    }

    [Fact]
    public void ParseList_ReadsCategoriesAndBundlesInOneListing()
    {
        var data = Category("Announcements", isBundle: false, articleCount: 12)
            .Concat(Category("Archive", isBundle: true, articleCount: 0))
            .Concat(Category("General", isBundle: false, articleCount: 3))
            .ToArray();

        var categories = HotlineNewsCategory.ParseList(data);

        Assert.Equal(3, categories.Count);
        Assert.Equal(new HotlineNewsCategory("Announcements", false, 12), categories[0]);
        Assert.Equal(new HotlineNewsCategory("Archive", true, 0), categories[1]);
        Assert.Equal(new HotlineNewsCategory("General", false, 3), categories[2]);
    }

    /// <summary>A listing cut short mid-entry should still yield the entries that were complete.</summary>
    [Fact]
    public void ParseList_TruncatedListing_KeepsWhatItCouldRead()
    {
        var full = Category("Announcements", isBundle: false, articleCount: 1)
            .Concat(Category("General", isBundle: false, articleCount: 2))
            .ToArray();

        var categories = HotlineNewsCategory.ParseList(full.AsSpan(0, full.Length - 4));

        Assert.Single(categories);
        Assert.Equal("Announcements", categories[0].Name);
    }

    private static byte[] ArticleList(params (uint Id, string Title, string Poster, uint Parent)[] articles)
    {
        var body = new List<byte>();
        body.AddRange(BigEndian(articles.Length == 0 ? 0u : articles[0].Id));
        body.AddRange(BigEndian((uint)articles.Length));
        body.AddRange([0, 0]); // name length: none

        foreach (var (id, title, poster, parent) in articles)
        {
            body.AddRange(BigEndian(id));
            body.AddRange(HotlineDate(2026, 3600));
            body.AddRange(BigEndian(parent));
            body.AddRange([0, 0]);       // flags
            body.AddRange([0, 1]);       // one flavour
            body.Add((byte)title.Length);
            body.AddRange(Encoding.UTF8.GetBytes(title));
            body.Add((byte)poster.Length);
            body.AddRange(Encoding.UTF8.GetBytes(poster));
            body.Add((byte)"text/plain".Length);
            body.AddRange("text/plain"u8.ToArray());
            body.AddRange([0, 16]);      // that flavour's size
        }

        return [.. body];
    }

    private static byte[] BigEndian(uint value)
    {
        var bytes = new byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(bytes, value);
        return bytes;
    }

    private static byte[] HotlineDate(ushort year, uint seconds)
    {
        var date = new byte[8];
        BinaryPrimitives.WriteUInt16BigEndian(date, year);
        BinaryPrimitives.WriteUInt32BigEndian(date.AsSpan(4), seconds);
        return date;
    }

    /// <summary>
    /// The flavour list is variable-length and sits at the end of each article, so a parser that
    /// skips it wrong lands mid-way through the next one. Two articles is the minimum that catches
    /// that.
    /// </summary>
    [Fact]
    public void ParseList_ReadsEveryArticle_IncludingPastTheVariableLengthFlavourList()
    {
        var data = ArticleList(
            (10, "Welcome", "Admin", 0),
            (11, "Re: Welcome", "Guest", 10));

        var articles = HotlineNewsArticle.ParseList(data);

        Assert.Equal(2, articles.Count);
        Assert.Equal(10u, articles[0].Id);
        Assert.Equal("Welcome", articles[0].Title);
        Assert.Equal("Admin", articles[0].Poster);
        Assert.Equal(0u, articles[0].ParentId);
        Assert.Equal("text/plain", articles[0].Flavor);

        Assert.Equal(11u, articles[1].Id);
        Assert.Equal("Re: Welcome", articles[1].Title);
        Assert.Equal(10u, articles[1].ParentId);
    }

    [Fact]
    public void ParseList_ReadsThePostedDate()
    {
        var articles = HotlineNewsArticle.ParseList(ArticleList((1, "T", "P", 0)));

        Assert.Equal(new DateTimeOffset(2026, 1, 1, 1, 0, 0, TimeSpan.Zero), articles[0].Posted);
    }

    [Fact]
    public void ParseList_EmptyCategory_YieldsNothing() => Assert.Empty(HotlineNewsArticle.ParseList(ArticleList()));
}

public class HotlineFileEntryTests
{
    private static byte[] Entry(string name, string type, uint size)
    {
        var nameBytes = Encoding.UTF8.GetBytes(name);
        var data = new byte[22 + nameBytes.Length];
        Encoding.ASCII.GetBytes(type).CopyTo(data, 0);
        Encoding.ASCII.GetBytes("HTLC").CopyTo(data, 4);
        BinaryPrimitives.WriteUInt32BigEndian(data.AsSpan(8), size);
        BinaryPrimitives.WriteUInt16BigEndian(data.AsSpan(20), (ushort)nameBytes.Length);
        nameBytes.CopyTo(data.AsSpan(22));
        return data;
    }

    [Fact]
    public void TryParse_ReadsAFile()
    {
        var entry = HotlineFileEntry.TryParse(Entry("readme.txt", "TEXT", 2048));

        Assert.NotNull(entry);
        Assert.Equal("readme.txt", entry.Name);
        Assert.Equal("TEXT", entry.TypeCode);
        Assert.Equal(2048u, entry.Size);
        Assert.False(entry.IsFolder);
    }

    /// <summary>A folder is marked by its type code, and its "size" is an item count — not bytes.</summary>
    [Fact]
    public void TryParse_ReadsAFolder()
    {
        var entry = HotlineFileEntry.TryParse(Entry("Uploads", "fldr", 7));

        Assert.NotNull(entry);
        Assert.True(entry.IsFolder);
        Assert.Equal(7u, entry.Size);
    }

    [Fact]
    public void TryParse_TooShortOrInconsistent_IsSkippedRatherThanThrowing()
    {
        Assert.Null(HotlineFileEntry.TryParse(new byte[10]));

        var truncated = Entry("readme.txt", "TEXT", 1)[..24];
        Assert.Null(HotlineFileEntry.TryParse(truncated));
    }
}

/// <summary>
/// The transfer connection, against a loopback stand-in for the server. These cover the framing a
/// real server sees and sends — the HTXF header, and a flattened file whose DATA fork is what a
/// download must actually keep.
/// </summary>
public class HotlineFileTransferTests
{
    [Fact]
    public void TransferPort_IsOneAboveTheServerPort() => Assert.Equal(5501, HotlineFileTransfer.TransferPort(5500));

    private static byte[] FlatFile(byte[] contents, bool withInfoFork = true)
    {
        var body = new List<byte>();
        body.AddRange("FILP"u8.ToArray());
        body.AddRange(new byte[18]);
        body.AddRange([0, (byte)(withInfoFork ? 2 : 1)]);

        if (withInfoFork)
        {
            var info = new byte[40];
            body.AddRange(ForkHeader("INFO", (uint)info.Length));
            body.AddRange(info);
        }

        body.AddRange(ForkHeader("DATA", (uint)contents.Length));
        body.AddRange(contents);
        return [.. body];
    }

    private static byte[] ForkHeader(string type, uint size)
    {
        var header = new byte[16];
        Encoding.ASCII.GetBytes(type).CopyTo(header, 0);
        BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(12), size);
        return header;
    }

    /// <summary>The DATA fork is the file; INFO must be read past, not written to disk.</summary>
    [Fact]
    public async Task Download_KeepsOnlyTheDataFork()
    {
        var contents = "the actual file contents"u8.ToArray();
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;

        byte[] headerSeen = new byte[16];
        var served = Task.Run(async () =>
        {
            using var client = await listener.AcceptTcpClientAsync();
            await using var stream = client.GetStream();
            await stream.ReadExactlyAsync(headerSeen);
            await stream.WriteAsync(FlatFile(contents));
        });

        using var destination = new MemoryStream();
        // The transfer port is serverPort+1, so the "server port" told to the API is one below.
        var written = await HotlineFileTransfer.DownloadAsync("127.0.0.1", port - 1, 0xABCD, destination);
        await served;

        Assert.Equal(contents, destination.ToArray());
        Assert.Equal(contents.Length, written);
        Assert.Equal("HTXF", Encoding.ASCII.GetString(headerSeen, 0, 4));
        Assert.Equal(0xABCDu, BinaryPrimitives.ReadUInt32BigEndian(headerSeen.AsSpan(4)));
        Assert.Equal(0u, BinaryPrimitives.ReadUInt32BigEndian(headerSeen.AsSpan(8)));
    }

    [Fact]
    public async Task Download_FileWithNoInfoFork_StillWorks()
    {
        var contents = "bare"u8.ToArray();
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;

        var served = Task.Run(async () =>
        {
            using var client = await listener.AcceptTcpClientAsync();
            await using var stream = client.GetStream();
            await stream.ReadExactlyAsync(new byte[16]);
            await stream.WriteAsync(FlatFile(contents, withInfoFork: false));
        });

        using var destination = new MemoryStream();
        await HotlineFileTransfer.DownloadAsync("127.0.0.1", port - 1, 1, destination);
        await served;

        Assert.Equal(contents, destination.ToArray());
    }

    /// <summary>A transfer that dies mid-file must fail loudly rather than leaving a short file that looks complete.</summary>
    [Fact]
    public async Task Download_ThatEndsEarly_Throws()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;

        _ = Task.Run(async () =>
        {
            using var client = await listener.AcceptTcpClientAsync();
            await using var stream = client.GetStream();
            await stream.ReadExactlyAsync(new byte[16]);
            // Announces 1000 bytes of DATA, sends 4, then hangs up.
            var truncated = FlatFile(new byte[1000], withInfoFork: false).AsSpan(0, 24 + 16 + 4).ToArray();
            await stream.WriteAsync(truncated);
        });

        using var destination = new MemoryStream();

        await Assert.ThrowsAnyAsync<IOException>(() =>
            HotlineFileTransfer.DownloadAsync("127.0.0.1", port - 1, 1, destination));
    }

    [Fact]
    public async Task Download_SomethingThatIsntAHotlineFile_Throws()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;

        _ = Task.Run(async () =>
        {
            using var client = await listener.AcceptTcpClientAsync();
            await using var stream = client.GetStream();
            await stream.ReadExactlyAsync(new byte[16]);
            await stream.WriteAsync(new byte[24]); // no FILP magic
        });

        using var destination = new MemoryStream();

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            HotlineFileTransfer.DownloadAsync("127.0.0.1", port - 1, 1, destination));
    }

    /// <summary>An upload announces its total size up front, then sends INFO and DATA — the server reads the size from the HTXF header to know when it's done.</summary>
    [Fact]
    public async Task Upload_SendsTheHeaderThenAFlattenedFile()
    {
        var contents = "uploaded bytes"u8.ToArray();
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;

        var received = new MemoryStream();
        var served = Task.Run(async () =>
        {
            using var client = await listener.AcceptTcpClientAsync();
            await using var stream = client.GetStream();
            await stream.CopyToAsync(received);
        });

        using var source = new MemoryStream(contents);
        await HotlineFileTransfer.UploadAsync("127.0.0.1", port - 1, 0x1234, "notes.txt", source, contents.Length);
        await served.WaitAsync(TimeSpan.FromSeconds(10));

        var sent = received.ToArray();
        Assert.Equal("HTXF", Encoding.ASCII.GetString(sent, 0, 4));
        Assert.Equal(0x1234u, BinaryPrimitives.ReadUInt32BigEndian(sent.AsSpan(4)));

        // The announced size must match what actually followed the 16-byte header.
        Assert.Equal((uint)(sent.Length - 16), BinaryPrimitives.ReadUInt32BigEndian(sent.AsSpan(8)));
        Assert.Equal("FILP", Encoding.ASCII.GetString(sent, 16, 4));
        Assert.Equal(2, BinaryPrimitives.ReadUInt16BigEndian(sent.AsSpan(38)));
        Assert.EndsWith("uploaded bytes", Encoding.UTF8.GetString(sent));
        Assert.Contains("notes.txt", Encoding.UTF8.GetString(sent));
    }
}
