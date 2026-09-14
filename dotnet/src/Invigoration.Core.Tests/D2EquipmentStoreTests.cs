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

    // Only us.bnet.cc, for now: no other server (PvPGN, Atlas...) is ever sent a request for these files.
    [Fact]
    public void TrustedServers_IsOnlyUsBnetCc()
    {
        Assert.Equal([("us.bnet.cc", 6112)], D2EquipmentStore.TrustedServers);
        Assert.True(D2EquipmentStore.IsTrustedServer(" US.bnet.cc "));
        Assert.False(D2EquipmentStore.IsTrustedServer("127.0.0.1"));
        Assert.False(D2EquipmentStore.IsTrustedServer("pvpgn.bnetdocs.org"));
        Assert.False(D2EquipmentStore.IsTrustedServer("atlas.bnetdocs.org"));
        Assert.False(D2EquipmentStore.IsTrustedServer("useast.battle.net"));
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

[Collection("D2EquipmentStore")]
public class D2DataUpdateTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "invig-d2update-" + Guid.NewGuid().ToString("N"));

    public D2DataUpdateTests()
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

    private static byte[] Zip()
    {
        using var ms = new MemoryStream();
        using (var archive = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
        {
            using var writer = new StreamWriter(archive.CreateEntry("manifest.json").Open());
            writer.Write("{}");
        }

        return ms.ToArray();
    }

    [Fact]
    public void BeforeOptingIn_NothingIsEverNewer()
    {
        Assert.False(D2EquipmentStore.OptedIn);
        Assert.False(D2EquipmentStore.IsNewerOnServer(D2EquipmentStore.EquipmentFileName, 999));
    }

    [Fact]
    public void AfterOptingIn_OnlyAStrictlyNewerServerCopyCounts()
    {
        D2EquipmentStore.Save(Bytes, fileTime: 500);

        Assert.True(D2EquipmentStore.IsNewerOnServer(D2EquipmentStore.EquipmentFileName, 501));
        Assert.False(D2EquipmentStore.IsNewerOnServer(D2EquipmentStore.EquipmentFileName, 500));
        Assert.False(D2EquipmentStore.IsNewerOnServer(D2EquipmentStore.EquipmentFileName, 0));
        Assert.False(D2EquipmentStore.IsNewerOnServer("icons.bni", 999));
    }

    // The art pack wasn't on the server when the user opted in; once it is, it's fetched without asking again.
    [Fact]
    public void AnArtPackWeDontHaveYet_CountsAsNewer()
    {
        D2EquipmentStore.Save(Bytes, fileTime: 500);

        Assert.True(D2EquipmentStore.IsNewerOnServer(D2EquipmentStore.CharacterPackFileName, 100));
    }

    [Fact]
    public void FileTimes_SurviveARestart()
    {
        D2EquipmentStore.Save(Bytes, fileTime: 123456789);
        D2EquipmentStore.SaveCharacterPack(Zip(), fileTime: 987654321);
        D2EquipmentStore.ResetCacheForTests();

        Assert.Equal(123456789, D2EquipmentStore.StoredFileTime(D2EquipmentStore.EquipmentFileName));
        Assert.Equal(987654321, D2EquipmentStore.StoredFileTime(D2EquipmentStore.CharacterPackFileName));
        Assert.True(D2EquipmentStore.HasCharacterPack);
    }

    [Fact]
    public void SaveCharacterPack_RejectsSomethingThatIsntAZip()
    {
        Assert.Throws<FormatException>(() => D2EquipmentStore.SaveCharacterPack("nope"u8.ToArray()));
        Assert.False(D2EquipmentStore.HasCharacterPack);
    }

    [Fact(Timeout = 15000)]
    public async Task DownloadAsync_FetchesTheArtPackToo_WhenTheServerHasIt()
    {
        var (port, _) = await ServeManyAsync(name => name == D2EquipmentStore.EquipmentFileName ? Bytes : Zip(), connections: 2);

        var (result, detail) = await D2EquipmentStore.DownloadAsync("127.0.0.1", port, TimeSpan.FromSeconds(5));

        Assert.Equal(D2EquipmentDownloadResult.Saved, result);
        Assert.True(D2EquipmentStore.HasCharacterPack);
        Assert.Contains("character art", detail);
    }

    [Fact(Timeout = 15000)]
    public async Task DownloadAsync_WithoutAnArtPack_StillSavesTheMap()
    {
        var (port, _) = await ServeManyAsync(name => name == D2EquipmentStore.EquipmentFileName ? Bytes : null, connections: 2);

        var (result, _) = await D2EquipmentStore.DownloadAsync("127.0.0.1", port, TimeSpan.FromSeconds(5));

        Assert.Equal(D2EquipmentDownloadResult.Saved, result);
        Assert.NotNull(D2EquipmentStore.Current);
        Assert.False(D2EquipmentStore.HasCharacterPack);
    }

    [Theory]
    [InlineData("us.bnet.cc", true, true)]
    [InlineData("US.BNET.CC", true, true)]
    [InlineData("us.bnet.cc", false, false)]
    [InlineData("127.0.0.1", true, false)]
    [InlineData("pvpgn.bnetdocs.org", true, false)]
    [InlineData("useast.battle.net", true, false)]
    public void TheEngineOnlyEverAsksATrustedServer_AfterOptingIn(string server, bool optedIn, bool expected)
    {
        var check = (bool)typeof(BotEngine).GetMethod("ShouldCheckD2FileTimes", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!
            .Invoke(null, [server, optedIn])!;

        Assert.Equal(expected, check);
    }
}

