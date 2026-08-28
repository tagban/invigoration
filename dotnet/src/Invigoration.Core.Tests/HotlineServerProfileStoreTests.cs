using Invigoration.Core.Hotline;

namespace Invigoration.Core.Tests;

/// <summary>Redirects HotlineServerProfileStore to an isolated temp directory for this test collection — same reasoning as BattlenetCredentialProfileStoreFixture: a hardcoded real-%AppData% path would otherwise leak permanent junk profiles into the user's actual config on every test run.</summary>
public sealed class HotlineServerProfileStoreFixture : IDisposable
{
    private readonly string _tempDir = Path.Combine(Path.GetTempPath(), $"invigoration-test-hotlineprofiles-{Guid.NewGuid():N}");

    public HotlineServerProfileStoreFixture() => HotlineServerProfileStore.ConfigDirectoryOverride = _tempDir;

    public void Dispose()
    {
        HotlineServerProfileStore.ConfigDirectoryOverride = null;
        try
        {
            Directory.Delete(_tempDir, recursive: true);
        }
        catch (DirectoryNotFoundException)
        {
        }
    }
}

[CollectionDefinition("HotlineServerProfileStore")]
public class HotlineServerProfileStoreCollection : ICollectionFixture<HotlineServerProfileStoreFixture>;

[Collection("HotlineServerProfileStore")]
public class HotlineServerProfileStoreTests
{
    [Fact]
    public void CreateAndSave_AssignsIdAndPersists()
    {
        var profile = HotlineServerProfileStore.CreateAndSave("Test Server", "hotline.example.com", 5500);

        Assert.False(string.IsNullOrEmpty(profile.Id));
        Assert.Equal("Test Server", profile.Name);
        Assert.Equal("hotline.example.com", profile.Host);
        Assert.Equal((ushort)5500, profile.Port);
        Assert.Contains(HotlineServerProfileStore.Profiles, p => p.Id == profile.Id);
    }

    [Fact]
    public void Find_UnknownId_ReturnsNull()
    {
        Assert.Null(HotlineServerProfileStore.Find($"nonexistent-{Guid.NewGuid():N}"));
    }

    [Fact]
    public void Delete_RemovesProfile()
    {
        var profile = HotlineServerProfileStore.CreateAndSave("To Delete", "host", 5500);

        HotlineServerProfileStore.Delete(profile.Id);

        Assert.DoesNotContain(HotlineServerProfileStore.Profiles, p => p.Id == profile.Id);
    }

    [Fact]
    public void Save_PersistsAcrossCacheReload()
    {
        var profile = HotlineServerProfileStore.CreateAndSave("Persisted", "host2", 5501);
        profile.AutoConnect = true;
        HotlineServerProfileStore.Save();

        var reloaded = JsonReload();

        Assert.Contains(reloaded, p => p.Id == profile.Id && p.AutoConnect);
    }

    private static List<HotlineServerProfile> JsonReload()
    {
        var json = File.ReadAllText(HotlineServerProfileStore.FilePath);
        return System.Text.Json.JsonSerializer.Deserialize<List<HotlineServerProfile>>(json) ?? [];
    }

    [Fact]
    public void NewProfile_HasDefaultPortAndIcon()
    {
        var profile = new HotlineServerProfile();

        Assert.Equal(HotlineConstants.DefaultServerPort, profile.Port);
        Assert.False(profile.AutoConnect);
    }
}
