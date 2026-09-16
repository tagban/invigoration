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

/// <summary>
/// The servers a brand-new install starts with. What they must NOT carry matters more than what
/// they do: anything that dials out on startup, or that agrees to a server's rules, or that
/// carries one person's own credentials, would be shipped to everybody.
/// </summary>
[Collection("HotlineServerProfileStore")]
public class HotlineDefaultProfileTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "invig-hlseed-" + Guid.NewGuid().ToString("N"));

    /// <summary>
    /// What the override was before this class touched it — restored on the way out rather than
    /// cleared. Clearing it is what let a test write to the real %AppData% profile list: this
    /// collection's fixture sets one shared override for every class in it, and a Dispose that
    /// nulled the property instead of putting the previous value back left every test that ran
    /// afterwards pointed at the user's own config. It really did create profiles there.
    /// </summary>
    private readonly string? _previousOverride = HotlineServerProfileStore.ConfigDirectoryOverride;

    public HotlineDefaultProfileTests()
    {
        Directory.CreateDirectory(_directory);
        HotlineServerProfileStore.ConfigDirectoryOverride = _directory;
    }

    public void Dispose()
    {
        HotlineServerProfileStore.ConfigDirectoryOverride = _previousOverride;
        try
        {
            Directory.Delete(_directory, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    [Fact]
    public void AFreshInstall_StartsWithTheDefaultServers()
    {
        var profiles = HotlineServerProfileStore.Profiles;

        Assert.Contains(profiles, p => p.Name == "MacDomain");
        Assert.Contains(profiles, p => p.Name == "HL Central");
    }

    [Fact]
    public void NoDefaultServer_ConnectsOnItsOwnOrCarriesCredentials()
    {
        foreach (var profile in HotlineServerProfileStore.DefaultProfiles())
        {
            Assert.False(profile.AutoConnect, $"{profile.Name} would connect on startup");
            Assert.False(profile.AutoAcceptAgreement, $"{profile.Name} would agree to the server's rules unasked");
            Assert.Equal("", profile.Login);
            Assert.Equal("", profile.Password);
            Assert.Equal("Guest", profile.Nickname);
        }
    }

    /// <summary>These servers bridge their chat to Discord under a specific account; without the name, relayed messages look like an ordinary user talking.</summary>
    [Fact]
    public void TheDefaultServers_KeepTheirDiscordRelayNames()
    {
        var defaults = HotlineServerProfileStore.DefaultProfiles();

        Assert.Equal("Discord", defaults.Single(p => p.Name == "MacDomain").DiscordRelayUsername);
        Assert.Equal("Relay", defaults.Single(p => p.Name == "HL Central").DiscordRelayUsername);
    }

    /// <summary>Deleting them is meant to stick — the seed is keyed on the file's absence, not on the list being empty.</summary>
    [Fact]
    public void DeletingEveryDefault_DoesNotBringThemBack()
    {
        foreach (var profile in HotlineServerProfileStore.Profiles.ToList())
        {
            HotlineServerProfileStore.Delete(profile.Id);
        }

        // Force a reload from disk the way a restart would (setting the override drops the cache).
        HotlineServerProfileStore.ConfigDirectoryOverride = _directory;

        Assert.Empty(HotlineServerProfileStore.Profiles);
    }
}
