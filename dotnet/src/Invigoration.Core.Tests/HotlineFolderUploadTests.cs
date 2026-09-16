using System.Buffers.Binary;
using System.Net;
using System.Net.Sockets;
using System.Text;
using Invigoration.Core.Hotline;

namespace Invigoration.Core.Tests;

/// <summary>
/// Folder uploads, against a loopback stand-in playing the server. The server drives this exchange
/// — it asks for the next item, then decides whether it wants that item's contents — so the risk is
/// in the turn-taking: answering out of step desynchronizes everything after it.
/// </summary>
public class HotlineFolderUploadTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "invig-upload-" + Guid.NewGuid().ToString("N"));

    public HotlineFolderUploadTests()
    {
        Directory.CreateDirectory(Path.Combine(_directory, "Art"));
        File.WriteAllText(Path.Combine(_directory, "readme.txt"), "read me");
        File.WriteAllText(Path.Combine(_directory, "Art", "logo.txt"), "the logo");
    }

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

    [Fact]
    public void Enumerate_ListsFoldersBeforeTheFilesInsideThem()
    {
        var items = HotlineFolderTransfer.Enumerate(_directory);

        // The folder has to exist on the server before the file that goes in it is sent.
        var folderIndex = items.ToList().FindIndex(i => i.IsFolder && i.Path[^1] == "Art");
        var fileIndex = items.ToList().FindIndex(i => !i.IsFolder && i.Path.Count == 2);

        Assert.True(folderIndex >= 0 && fileIndex > folderIndex);
        Assert.Equal(3, items.Count);
        Assert.Equal(7, items.Single(i => i.Path[^1] == "readme.txt").Size);
    }

    private static async Task<ushort> ReadUInt16Async(Stream stream)
    {
        var buffer = new byte[2];
        await stream.ReadExactlyAsync(buffer);
        return BinaryPrimitives.ReadUInt16BigEndian(buffer);
    }

    private static async Task SendActionAsync(Stream stream, ushort action)
    {
        var bytes = new byte[2];
        BinaryPrimitives.WriteUInt16BigEndian(bytes, action);
        await stream.WriteAsync(bytes);
    }

    /// <summary>Reads one item descriptor and returns whether it's a folder plus its path components.</summary>
    private static async Task<(bool IsFolder, List<string> Path)> ReadItemAsync(Stream stream)
    {
        var size = await ReadUInt16Async(stream);
        var body = new byte[size];
        await stream.ReadExactlyAsync(body);

        var isFolder = BinaryPrimitives.ReadUInt16BigEndian(body) == 1;
        var count = BinaryPrimitives.ReadUInt16BigEndian(body.AsSpan(2));
        var path = new List<string>();
        var offset = 4;
        for (var i = 0; i < count; i++)
        {
            var length = body[offset + 2];
            path.Add(Encoding.UTF8.GetString(body, offset + 3, length));
            offset += 3 + length;
        }

        return (isFolder, path);
    }

    [Fact]
    public async Task Upload_OffersEveryItemAndSendsTheFilesTheServerAsksFor()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;

        var offered = new List<(bool IsFolder, List<string> Path)>();
        var contents = new List<string>();
        var openingHeader = new byte[16];

        var served = Task.Run(async () =>
        {
            using var client = await listener.AcceptTcpClientAsync();
            await using var stream = client.GetStream();
            await stream.ReadExactlyAsync(openingHeader);

            for (var i = 0; i < 3; i++)
            {
                await SendActionAsync(stream, 3); // next file
                var item = await ReadItemAsync(stream);
                offered.Add(item);

                if (item.IsFolder)
                {
                    continue;
                }

                await SendActionAsync(stream, 1); // send it

                var sizeBytes = new byte[4];
                await stream.ReadExactlyAsync(sizeBytes);
                var flattened = new byte[BinaryPrimitives.ReadUInt32BigEndian(sizeBytes)];
                await stream.ReadExactlyAsync(flattened);

                // The DATA fork is the last thing in the flattened file.
                var dataIndex = flattened.AsSpan().LastIndexOf("DATA"u8);
                var dataSize = (int)BinaryPrimitives.ReadUInt32BigEndian(flattened.AsSpan(dataIndex + 12));
                contents.Add(Encoding.UTF8.GetString(flattened, dataIndex + 16, dataSize));
            }
        });

        var sent = await HotlineFolderTransfer.UploadAsync("127.0.0.1", port - 1, 0xCAFE, _directory);
        await served.WaitAsync(TimeSpan.FromSeconds(10));

        Assert.Equal(2, sent);
        Assert.Equal(3, offered.Count);
        Assert.Contains(offered, o => o.IsFolder && o.Path is ["Art"]);
        Assert.Contains(offered, o => !o.IsFolder && o.Path is ["Art", "logo.txt"]);
        Assert.Contains("the logo", contents);
        Assert.Contains("read me", contents);

        Assert.Equal("HTXF", Encoding.ASCII.GetString(openingHeader, 0, 4));
        Assert.Equal(0xCAFEu, BinaryPrimitives.ReadUInt32BigEndian(openingHeader.AsSpan(4)));
        Assert.Equal(1, BinaryPrimitives.ReadUInt16BigEndian(openingHeader.AsSpan(12))); // folder transfer
    }

    /// <summary>A server that already has a file says so; the client must move on rather than sending it anyway.</summary>
    [Fact]
    public async Task Upload_AnItemTheServerSkips_IsNotSent()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;

        var served = Task.Run(async () =>
        {
            using var client = await listener.AcceptTcpClientAsync();
            await using var stream = client.GetStream();
            await stream.ReadExactlyAsync(new byte[16]);

            for (var i = 0; i < 3; i++)
            {
                await SendActionAsync(stream, 3);
                var item = await ReadItemAsync(stream);
                if (!item.IsFolder)
                {
                    await SendActionAsync(stream, 3); // skip this one
                }
            }
        });

        var sent = await HotlineFolderTransfer.UploadAsync("127.0.0.1", port - 1, 1, _directory);
        await served.WaitAsync(TimeSpan.FromSeconds(10));

        Assert.Equal(0, sent);
    }
}
