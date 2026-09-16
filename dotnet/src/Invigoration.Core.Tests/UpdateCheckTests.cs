using Invigoration.Core.Updates;

namespace Invigoration.Core.Tests;

public class ReleaseVersionTests
{
    [Theory]
    [InlineData("2.0.8b", "2.0.7b")]   // patch
    [InlineData("2.0.10b", "2.0.9b")]  // two digits — the case string comparison gets wrong
    [InlineData("2.1.0b", "2.0.99b")]
    [InlineData("3.0.0b", "2.9.9b")]
    [InlineData("2.0.8b", "2.0.8a")]   // same numbers, later letter
    [InlineData("2.0.8", "2.0.7b")]
    [InlineData("2.1.0", "2.0.8b")]    // leaving beta: the number has to go up, or nobody is offered it
    [InlineData("2.1.0", "2.0.9")]
    public void IsNewerThan_RecognizesTheLaterVersion(string later, string earlier)
    {
        var a = ReleaseVersion.TryParse(later)!;
        var b = ReleaseVersion.TryParse(earlier)!;

        Assert.True(a.IsNewerThan(b));
        Assert.False(b.IsNewerThan(a));
    }

    [Theory]
    [InlineData("2.0.8b", "v2.0.8b")]  // release tags carry a "v"
    [InlineData("2.0.8b", "2.0.8b")]
    [InlineData("2.1", "2.1.0")]       // a missing part is zero
    public void EquivalentSpellings_AreNeitherNewerNorOlder(string one, string other)
    {
        var a = ReleaseVersion.TryParse(one)!;
        var b = ReleaseVersion.TryParse(other)!;

        Assert.False(a.IsNewerThan(b));
        Assert.False(b.IsNewerThan(a));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    [InlineData("nightly")]
    [InlineData("v")]
    [InlineData("2.x.0")]
    public void Nonsense_ParsesToNothing(string? text) => Assert.Null(ReleaseVersion.TryParse(text));

    /// <summary>
    /// The trap that decided 2.1.0's number: a bare "2.0.8" sorts BELOW "2.0.8b", because the
    /// suffix is compared as text and "" comes before "b". Dropping the beta letter without
    /// raising the number would have left everyone on 2.0.8b never offered the stable build.
    /// </summary>
    [Fact]
    public void DroppingTheBetaLetterAlone_WouldLookOlderNotNewer()
    {
        var stable = ReleaseVersion.TryParse("2.0.8")!;
        var beta = ReleaseVersion.TryParse("2.0.8b")!;

        Assert.False(stable.IsNewerThan(beta));
        Assert.True(beta.IsNewerThan(stable));

        // Which is why the release went out as 2.1.0.
        Assert.True(ReleaseVersion.TryParse("2.1.0")!.IsNewerThan(beta));
    }

    [Fact]
    public void ToString_RoundTripsWithoutTheTagPrefix()
    {
        Assert.Equal("2.0.8b", ReleaseVersion.TryParse("v2.0.8b")!.ToString());
    }
}

/// <summary>
/// The update notice is meant to be easy to ignore and impossible to be bothered by twice, so
/// these pin the four ways it must stay quiet: checking is off, nothing newer exists, the version
/// was already dismissed, or the check failed.
/// </summary>
[Collection("UpdateSettings")]
public class UpdateCheckTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "invig-update-" + Guid.NewGuid().ToString("N"));

    /// <summary>Restored rather than cleared on the way out — clearing a store's directory override points whatever runs next at the user's real config. See HotlineDefaultProfileTests for the time that actually happened.</summary>
    private readonly string? _previousOverride = UpdateSettingsStore.DirectoryOverride;

    public UpdateCheckTests()
    {
        Directory.CreateDirectory(_directory);
        UpdateSettingsStore.DirectoryOverride = _directory;
        UpdateSettingsStore.ResetCacheForTests();
    }

    public void Dispose()
    {
        UpdateCheck.LatestTagOverride = null;
        UpdateSettingsStore.DirectoryOverride = _previousOverride;
        UpdateSettingsStore.ResetCacheForTests();
        try
        {
            Directory.Delete(_directory, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    private static void ServerHas(string? tag) => UpdateCheck.LatestTagOverride = _ => Task.FromResult(tag);

    [Fact]
    public async Task ANewerRelease_IsOffered()
    {
        ServerHas("v2.0.9b");

        var update = await UpdateCheck.FindNewerReleaseAsync("2.0.8b");

        Assert.NotNull(update);
        Assert.Equal("2.0.9b", update.Version.ToString());
        Assert.Equal(UpdateCheck.ReleasesPage, update.ReleasesUrl);
    }

    [Theory]
    [InlineData("v2.0.8b")] // the same version
    [InlineData("v2.0.7b")] // older than what's running
    public async Task NothingNewer_SaysNothing(string tag)
    {
        ServerHas(tag);

        Assert.Null(await UpdateCheck.FindNewerReleaseAsync("2.0.8b"));
    }

    [Fact]
    public async Task ADismissedVersion_IsNotOfferedAgain()
    {
        ServerHas("v2.0.9b");
        UpdateSettingsStore.DismissedVersion = "2.0.9b";

        Assert.Null(await UpdateCheck.FindNewerReleaseAsync("2.0.8b"));
    }

    /// <summary>Dismissing one version must not mute every later one — that would turn "not now" into "never".</summary>
    [Fact]
    public async Task AReleaseNewerThanTheDismissedOne_IsStillOffered()
    {
        UpdateSettingsStore.DismissedVersion = "2.0.9b";
        ServerHas("v2.1.0b");

        var update = await UpdateCheck.FindNewerReleaseAsync("2.0.8b");

        Assert.Equal("2.1.0b", update?.Version.ToString());
    }

    [Fact]
    public async Task WithCheckingTurnedOff_NothingIsAsked()
    {
        var asked = false;
        UpdateCheck.LatestTagOverride = _ =>
        {
            asked = true;
            return Task.FromResult<string?>("v9.9.9b");
        };
        UpdateSettingsStore.CheckForUpdates = false;

        Assert.Null(await UpdateCheck.FindNewerReleaseAsync("2.0.8b"));
        Assert.False(asked, "the check must not contact GitHub at all when it's turned off");
    }

    [Theory]
    [InlineData(null)]        // the request failed
    [InlineData("nightly")]   // a tag that isn't a version
    public async Task AnUnusableAnswer_IsIgnoredQuietly(string? tag)
    {
        ServerHas(tag);

        Assert.Null(await UpdateCheck.FindNewerReleaseAsync("2.0.8b"));
    }

    /// <summary>A local build ahead of the last release shouldn't be told to "update" backwards.</summary>
    [Fact]
    public async Task ABuildAheadOfTheLatestRelease_IsLeftAlone()
    {
        ServerHas("v2.0.8b");

        Assert.Null(await UpdateCheck.FindNewerReleaseAsync("2.0.9b"));
    }

    [Fact]
    public void CheckingIsOnByDefault()
    {
        Assert.True(UpdateSettingsStore.CheckForUpdates);
        Assert.Equal("", UpdateSettingsStore.DismissedVersion);
    }
}

/// <summary>UpdateSettingsStore is process-wide static state, so these classes run one at a time.</summary>
[CollectionDefinition("UpdateSettings")]
public class UpdateSettingsCollection;
