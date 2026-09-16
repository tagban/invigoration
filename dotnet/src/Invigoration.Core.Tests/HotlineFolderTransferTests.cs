using System.Buffers.Binary;
using System.Net;
using System.Net.Sockets;
using System.Text;
using Invigoration.Core.Hotline;

namespace Invigoration.Core.Tests;

/// <summary>
/// Folder downloads, against a loopback stand-in. Unlike a single file, this is a conversation:
/// the server describes an item, the client says send-it or skip-it, and that repeats. Getting the
/// action bytes or the item-header length wrong desynchronizes the whole exchange, so these drive
/// the real loop rather than checking pieces in isolation.
/// </summary>
public class HotlineFolderTransferTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "invig-folder-" + Guid.NewGuid().ToString("N"));

    public HotlineFolderTransferTests() => Directory.CreateDirectory(_directory);

    public void Dispose()
    {
        try
        {
            Directory.Delete(_directory, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    /// <summary>One item header: a size covering the type and path, the type, then the packed path.</summary>
    private static byte[] ItemHeader(IReadOnlyList<string> path, bool isFolder)
    {
        var packed = HotlineNewsPath.Encode(path);
        var header = new byte[4 + packed.Length];
        BinaryPrimitives.WriteUInt16BigEndian(header, (ushort)(2 + packed.Length));
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(2), isFolder ? (ushort)1 : (ushort)0);
        packed.CopyTo(header.AsSpan(4));
        return header;
    }

    private static byte[] FlatFile(byte[] contents)
    {
        var body = new List<byte>();
        body.AddRange("FILP"u8.ToArray());
        body.AddRange(new byte[18]);
        body.AddRange([0, 1]); // one fork: DATA
        var forkHeader = new byte[16];
        "DATA"u8.CopyTo(forkHeader.AsSpan(0));
        BinaryPrimitives.WriteUInt32BigEndian(forkHeader.AsSpan(12), (uint)contents.Length);
        body.AddRange(forkHeader);
        body.AddRange(contents);
        return [.. body];
    }

    private static byte[] SizedFlatFile(byte[] contents)
    {
        var flat = FlatFile(contents);
        var framed = new byte[4 + flat.Length];
        BinaryPrimitives.WriteUInt32BigEndian(framed, (uint)flat.Length);
        flat.CopyTo(framed.AsSpan(4));
        return framed;
    }

    private static async Task<ushort> ReadActionAsync(Stream stream)
    {
        var action = new byte[2];
        await stream.ReadExactlyAsync(action);
        return BinaryPrimitives.ReadUInt16BigEndian(action);
    }

    [Fact]
    public async Task Download_WalksEveryItem_RecreatingFoldersAndWritingFiles()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;

        var openingHeader = new byte[18];
        var actions = new List<ushort>();

        var served = Task.Run(async () =>
        {
            using var client = await listener.AcceptTcpClientAsync();
            await using var stream = client.GetStream();
            await stream.ReadExactlyAsync(openingHeader);

            // A folder, then a file inside it, then a file at the top level.
            await stream.WriteAsync(ItemHeader(["Art"], isFolder: true));
            actions.Add(await ReadActionAsync(stream));

            await stream.WriteAsync(ItemHeader(["Art", "logo.png"], isFolder: false));
            actions.Add(await ReadActionAsync(stream));
            await stream.WriteAsync(SizedFlatFile("the logo"u8.ToArray()));

            await stream.WriteAsync(ItemHeader(["readme.txt"], isFolder: false));
            actions.Add(await ReadActionAsync(stream));
            await stream.WriteAsync(SizedFlatFile("read me"u8.ToArray()));
        });

        var items = new List<HotlineFolderItem>();
        var written = await HotlineFolderTransfer.DownloadAsync(
            "127.0.0.1", port - 1, 0xBEEF, itemCount: 3, _directory,
            itemProgress: new Progress<HotlineFolderItem>(items.Add));
        await served.WaitAsync(TimeSpan.FromSeconds(10));

        Assert.Equal(2, written);
        Assert.True(Directory.Exists(Path.Combine(_directory, "Art")));
        Assert.Equal("the logo", await File.ReadAllTextAsync(Path.Combine(_directory, "Art", "logo.png")));
        Assert.Equal("read me", await File.ReadAllTextAsync(Path.Combine(_directory, "readme.txt")));

        // A folder is skipped past; each file is asked for.
        Assert.Equal([3, 1, 1], actions);

        // The opening header announces a folder transfer and starts the loop.
        Assert.Equal("HTXF", Encoding.ASCII.GetString(openingHeader, 0, 4));
        Assert.Equal(0xBEEFu, BinaryPrimitives.ReadUInt32BigEndian(openingHeader.AsSpan(4)));
        Assert.Equal(1, BinaryPrimitives.ReadUInt16BigEndian(openingHeader.AsSpan(12)));
        Assert.Equal(3, BinaryPrimitives.ReadUInt16BigEndian(openingHeader.AsSpan(16)));
    }

    /// <summary>A server that promises more items than it sends shouldn't hang the download forever.</summary>
    [Fact]
    public async Task Download_AServerThatStopsEarly_FinishesWithWhatItGot()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;

        _ = Task.Run(async () =>
        {
            using var client = await listener.AcceptTcpClientAsync();
            await using var stream = client.GetStream();
            await stream.ReadExactlyAsync(new byte[18]);
            await stream.WriteAsync(ItemHeader(["only.txt"], isFolder: false));
            await ReadActionAsync(stream);
            await stream.WriteAsync(SizedFlatFile("just one"u8.ToArray()));
            // Promised 5, hangs up after 1.
        });

        var written = await HotlineFolderTransfer.DownloadAsync("127.0.0.1", port - 1, 1, itemCount: 5, _directory);

        Assert.Equal(1, written);
        Assert.Equal("just one", await File.ReadAllTextAsync(Path.Combine(_directory, "only.txt")));
    }

    /// <summary>
    /// A server is not trusted to send well-behaved names. A path that climbs out of the chosen
    /// folder must be neutralized, or a folder download could write anywhere on disk.
    /// </summary>
    [Theory]
    [InlineData(new[] { "..", "escaped.txt" }, "escaped.txt")]
    [InlineData(new[] { ".", "dotted.txt" }, "dotted.txt")]
    [InlineData(new[] { "", "empty.txt" }, "empty.txt")]
    public void SafePath_StripsAnythingThatWouldEscapeTheFolder(string[] components, string expected) =>
        Assert.Equal([expected], HotlineFolderTransfer.SafePath(components).ToList());

    [Fact]
    public void SafePath_ReplacesCharactersAFilesystemWontTake()
    {
        var cleaned = HotlineFolderTransfer.SafePath(["a/b"]).Single();

        Assert.DoesNotContain(Path.GetInvalidFileNameChars(), cleaned.Contains);
    }

    /// <summary>The traversal guard has to hold for the real download, not just the helper.</summary>
    [Fact]
    public async Task Download_APathClimbingOutOfTheFolder_StaysInside()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;

        _ = Task.Run(async () =>
        {
            using var client = await listener.AcceptTcpClientAsync();
            await using var stream = client.GetStream();
            await stream.ReadExactlyAsync(new byte[18]);
            await stream.WriteAsync(ItemHeader(["..", "..", "owned.txt"], isFolder: false));
            await ReadActionAsync(stream);
            await stream.WriteAsync(SizedFlatFile("nope"u8.ToArray()));
        });

        await HotlineFolderTransfer.DownloadAsync("127.0.0.1", port - 1, 1, itemCount: 1, _directory);

        Assert.True(File.Exists(Path.Combine(_directory, "owned.txt")));
        Assert.False(File.Exists(Path.Combine(_directory, "..", "..", "owned.txt")));
    }
}
