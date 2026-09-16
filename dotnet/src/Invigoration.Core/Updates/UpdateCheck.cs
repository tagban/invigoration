using System.Net.Http.Json;
using System.Text.Json.Serialization;

namespace Invigoration.Core.Updates;

/// <summary>A newer release than the one running, and where to read about it.</summary>
public sealed record AvailableUpdate(ReleaseVersion Version, string ReleasesUrl);

/// <summary>
/// Asks GitHub once, at startup, whether there's a newer release than this build — nothing more.
/// It never downloads or installs anything, sends nothing about the user (no account, no machine
/// details, not even which version is running: the comparison happens here), and stays quiet on
/// any failure, since being offline or GitHub being slow is not something to interrupt anyone
/// about. The result is one dismissible line in the main window linking to the releases page.
/// Off entirely when <see cref="UpdateSettingsStore.CheckForUpdates"/> is false, which is the
/// only state in which this type contacts the network at all.
/// </summary>
public static class UpdateCheck
{
    public const string ReleasesPage = "https://github.com/tagban/invigoration/releases";
    private const string LatestReleaseApi = "https://api.github.com/repos/tagban/invigoration/releases/latest";

    /// <summary>Test hook: answers from here instead of GitHub.</summary>
    public static Func<CancellationToken, Task<string?>>? LatestTagOverride { get; set; }

    /// <summary>
    /// The newer release to tell the user about, or null — which covers every uninteresting case:
    /// checking is off, this build is current or newer (a local build ahead of the last release),
    /// the user already dismissed this version, or the check simply didn't work.
    /// </summary>
    public static async Task<AvailableUpdate?> FindNewerReleaseAsync(
        string currentVersion,
        CancellationToken cancellationToken = default)
    {
        if (!UpdateSettingsStore.CheckForUpdates)
        {
            return null;
        }

        if (ReleaseVersion.TryParse(currentVersion) is not { } current)
        {
            return null;
        }

        var tag = await TryGetLatestTagAsync(cancellationToken).ConfigureAwait(false);
        if (ReleaseVersion.TryParse(tag) is not { } latest || !latest.IsNewerThan(current))
        {
            return null;
        }

        return UpdateSettingsStore.ShouldAnnounce(latest) ? new AvailableUpdate(latest, ReleasesPage) : null;
    }

    private static async Task<string?> TryGetLatestTagAsync(CancellationToken cancellationToken)
    {
        if (LatestTagOverride is { } stub)
        {
            return await stub(cancellationToken).ConfigureAwait(false);
        }

        try
        {
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
            // GitHub's API rejects requests without one; it identifies the app, not the user.
            http.DefaultRequestHeaders.UserAgent.ParseAdd($"Invigoration/{AppVersion.Current}");
            http.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
            var release = await http.GetFromJsonAsync<LatestRelease>(LatestReleaseApi, cancellationToken).ConfigureAwait(false);
            return release?.TagName;
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or System.Text.Json.JsonException or UriFormatException)
        {
            // Offline, rate-limited, slow, or something unexpected in the response — all the same
            // here: no notice this run, no error in anyone's face.
            return null;
        }
    }

    private sealed record LatestRelease([property: JsonPropertyName("tag_name")] string? TagName);
}
