using Invigoration.Core.Networking;
using static Invigoration.Core.Tests.D2EquipmentTestData;

namespace Invigoration.Core.Tests;

public class BnftpClientTests
{
    private static readonly TimeSpan Timeout = TimeSpan.FromSeconds(5);

    [Fact]
    public void BuildRequest_MatchesTheVersion1Layout()
    {
        var request = BnftpClient.BuildRequest("icons.bni");

        Assert.Equal(0x02, request[0]);
        Assert.Equal(request.Length - 1, BitConverter.ToUInt16(request, 1));
        Assert.Equal(0x100, BitConverter.ToUInt16(request, 3));
        Assert.Equal("68XIRATS", System.Text.Encoding.ASCII.GetString(request, 5, 8));
        Assert.Equal("icons.bni\0", System.Text.Encoding.ASCII.GetString(request, 33, request.Length - 33));
    }

    [Fact(Timeout = 10000)]
    public async Task DownloadAsync_FetchesTheNamedFile()
    {
        var (port, requested) = await ServeOnceAsync(_ => Bytes);

        var file = await BnftpClient.DownloadAsync("127.0.0.1", port, "d2-equipment.json", Timeout);

        Assert.Equal("d2-equipment.json", await requested);
        Assert.NotNull(file);
        Assert.Equal("d2-equipment.json", file.Name);
        Assert.Equal(Bytes, file.Data);
        Assert.Equal(134335450650000000L, file.FileTime);
    }

    [Fact(Timeout = 10000)]
    public async Task DownloadAsync_NoSuchFile_IsNull()
    {
        var (port, _) = await ServeOnceAsync(_ => null);

        Assert.Null(await BnftpClient.DownloadAsync("127.0.0.1", port, "nope.json", Timeout));
    }

    [Fact(Timeout = 10000)]
    public async Task DownloadAsync_RefusesAnAbsurdAnnouncedSize()
    {
        var (port, _) = await ServeOnceAsync(_ => [1, 2, 3], announcedSize: 64 * 1024 * 1024);

        await Assert.ThrowsAsync<IOException>(() => BnftpClient.DownloadAsync("127.0.0.1", port, "big.bin", Timeout));
    }

    [Fact(Timeout = 10000)]
    public async Task DownloadAsync_ATransferCutShort_Throws()
    {
        var (port, _) = await ServeOnceAsync(_ => Bytes, sendOnly: 100);

        await Assert.ThrowsAsync<IOException>(() => BnftpClient.DownloadAsync("127.0.0.1", port, "d2-equipment.json", Timeout));
    }
}
