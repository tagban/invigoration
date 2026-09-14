using Invigoration.Core.Config;
using static Invigoration.Core.Tests.D2EquipmentTestData;

namespace Invigoration.Core.Tests;

/// <summary>Serialized: D2EquipmentStore is static, with a process-wide cache and a folder on disk.</summary>
[Collection("D2EquipmentStore")]
public class D2EquipmentStoreTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "invig-d2equip-" + Guid.NewGuid().ToString("N"));

    public D2EquipmentStoreTests()
    {
        D2EquipmentStore.DirectoryOverride = _dir;
        D2EquipmentStore.ResetCacheForTests();
    }

    public void Dispose()
    {
        D2EquipmentStore.DirectoryOverride = null;
        D2EquipmentStore.ResetCacheForTests();
        try
        {
            Directory.Delete(_dir, recursive: true);
        }
        catch (DirectoryNotFoundException)
        {
        }
    }

    [Fact]
    public void NothingDownloaded_HasNoMap_AndHasntBeenDeclined()
    {
        Assert.Null(D2EquipmentStore.Current);
        Assert.False(D2EquipmentStore.Declined);
    }

    [Fact]
    public void Save_KeepsTheMap_AcrossRestarts_AndTellsListeners()
    {
        var changed = 0;
        void OnChanged() => changed++;
        D2EquipmentStore.Changed += OnChanged;
        try
        {
            D2EquipmentStore.Save(Bytes);
        }
        finally
        {
            D2EquipmentStore.Changed -= OnChanged;
        }

        Assert.Equal(1, changed);
        Assert.NotNull(D2EquipmentStore.Current);

        D2EquipmentStore.ResetCacheForTests();
        Assert.NotNull(D2EquipmentStore.Current);
    }

    [Fact]
    public void Save_SomethingUnreadable_WritesNothing()
    {
        Assert.Throws<FormatException>(() => D2EquipmentStore.Save("{\"format\":\"nope\"}"u8.ToArray()));

        Assert.False(File.Exists(D2EquipmentStore.FilePath));
        Assert.Null(D2EquipmentStore.Current);
    }

    [Fact]
    public void ACorruptedStoredCopy_ReadsAsNoMap()
    {
        Directory.CreateDirectory(_dir);
        File.WriteAllText(D2EquipmentStore.FilePath, "{ broken");

        Assert.Null(D2EquipmentStore.Current);
    }

    [Fact]
    public void Declined_IsRemembered()
    {
        D2EquipmentStore.Declined = true;
        D2EquipmentStore.ResetCacheForTests();

        Assert.True(D2EquipmentStore.Declined);
    }

    [Fact(Timeout = 10000)]
    public async Task DownloadAsync_SavesWhatTheServerSends()
    {
        var (port, requested) = await ServeOnceAsync(_ => Bytes);

        var (result, _) = await D2EquipmentStore.DownloadAsync("127.0.0.1", port, TimeSpan.FromSeconds(5));

        Assert.Equal(D2EquipmentDownloadResult.Saved, result);
        Assert.Equal("d2-equipment.json", await requested);
        Assert.NotNull(D2EquipmentStore.Current);
    }

    [Fact(Timeout = 10000)]
    public async Task DownloadAsync_ServerWithoutTheFile()
    {
        var (port, _) = await ServeOnceAsync(_ => null);

        var (result, _) = await D2EquipmentStore.DownloadAsync("127.0.0.1", port, TimeSpan.FromSeconds(5));

        Assert.Equal(D2EquipmentDownloadResult.NotOnServer, result);
        Assert.Null(D2EquipmentStore.Current);
    }

    [Fact(Timeout = 10000)]
    public async Task DownloadAsync_ServerSendsJunk_KeepsNothing()
    {
        var (port, _) = await ServeOnceAsync(_ => "<html>not it</html>"u8.ToArray());

        var (result, _) = await D2EquipmentStore.DownloadAsync("127.0.0.1", port, TimeSpan.FromSeconds(5));

        Assert.Equal(D2EquipmentDownloadResult.Unreadable, result);
        Assert.Null(D2EquipmentStore.Current);
    }

    [Fact(Timeout = 10000)]
    public async Task DownloadAsync_NobodyListening_Fails()
    {
        var listener = new System.Net.Sockets.TcpListener(System.Net.IPAddress.Loopback, 0);
        listener.Start();
        var port = ((System.Net.IPEndPoint)listener.LocalEndpoint).Port;
        listener.Stop();

        var (result, _) = await D2EquipmentStore.DownloadAsync("127.0.0.1", port, TimeSpan.FromSeconds(5));

        Assert.Equal(D2EquipmentDownloadResult.Failed, result);
    }
}

[Collection("D2EquipmentStore")]
public class D2EquipmentStoreSourceTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "invig-d2equip-src-" + Guid.NewGuid().ToString("N"));

    public D2EquipmentStoreSourceTests()
    {
        D2EquipmentStore.DirectoryOverride = _dir;
        D2EquipmentStore.ResetCacheForTests();
    }

    public void Dispose()
    {
        D2EquipmentStore.DirectoryOverride = null;
        D2EquipmentStore.ResetCacheForTests();
        try
        {
            Directory.Delete(_dir, recursive: true);
        }
        catch (DirectoryNotFoundException)
        {
        }
    }

    [Fact]
    public void DownloadSources_AsksTheUsersOwnServersFirst_ThenUsBnetCc_NeverBlizzard()
    {
        var sources = D2EquipmentStore.DownloadSources(
        [
            new BotConfig { BattlenetServer = "useast.battle.net", BattlenetPort = 6112 },
            new BotConfig { BattlenetServer = "127.0.0.1", BattlenetPort = 6112 },
            new BotConfig { BattlenetServer = "127.0.0.1", BattlenetPort = 6112 },
            new BotConfig { BattlenetServer = "atlas.bnetdocs.org", BattlenetPort = 6112 },
            new BotConfig { BattlenetServer = "" },
        ]);

        Assert.Equal([("127.0.0.1", 6112), ("atlas.bnetdocs.org", 6112), (D2EquipmentStore.DefaultHost, 6112)], sources);
    }

    [Fact]
    public void DownloadSources_DoesntListUsBnetCcTwice()
    {
        var sources = D2EquipmentStore.DownloadSources([new BotConfig { BattlenetServer = "us.bnet.cc", BattlenetPort = 6112 }]);

        Assert.Equal([("us.bnet.cc", 6112)], sources);
    }

    [Fact(Timeout = 15000)]
    public async Task DownloadAsync_MovesOnPastServersWithoutIt()
    {
        var (missingPort, _) = await ServeOnceAsync(_ => null);
        var (servingPort, _) = await ServeOnceAsync(_ => Bytes);

        var (result, _) = await D2EquipmentStore.DownloadAsync([("127.0.0.1", missingPort), ("127.0.0.1", servingPort)], TimeSpan.FromSeconds(5));

        Assert.Equal(D2EquipmentDownloadResult.Saved, result);
        Assert.NotNull(D2EquipmentStore.Current);
    }

    [Fact(Timeout = 15000)]
    public async Task DownloadAsync_NobodyHasIt_SaysSo()
    {
        var (a, _) = await ServeOnceAsync(_ => null);
        var (b, _) = await ServeOnceAsync(_ => null);

        var (result, detail) = await D2EquipmentStore.DownloadAsync([("127.0.0.1", a), ("127.0.0.1", b)], TimeSpan.FromSeconds(5));

        Assert.Equal(D2EquipmentDownloadResult.NotOnServer, result);
        Assert.Contains("127.0.0.1", detail);
    }
}
