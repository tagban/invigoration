using System.Text.Json;
using Avalonia.Controls;
using Avalonia.Threading;
using Invigoration.Core.Music;

namespace Invigoration.App.Music;

/// <summary>
/// Drives a NativeWebView against whichever MusicServiceProfile is currently active, executing
/// JavaScript against the page's own DOM — the same technique the (unrelated)
/// th-ch/youtube-music desktop wrapper uses. Shared across services rather than one controller
/// class per service, since the only real difference between them is which selectors/scripts to
/// run, not how the run-a-script-and-report-success plumbing works.
/// </summary>
public sealed class WebViewMusicController(NativeWebView webView) : IMusicPlayerController
{
    public MusicServiceProfile Profile { get; set; } = MusicServiceProfile.YouTubeMusic;

    public Task<bool> SkipAsync() => ClickAsync(Profile.NextScript);

    public Task<bool> ThumbsUpAsync() => ClickAsync(Profile.LikeScript);

    public Task<bool> ThumbsDownAsync() => ClickAsync(Profile.DislikeScript);

    public bool SupportsThumbsUp => Profile.LikeScript is not null;

    public bool SupportsThumbsDown => Profile.DislikeScript is not null;

    private async Task<bool> ClickAsync(string? elementExpression)
    {
        if (elementExpression is null)
        {
            // e.g. Spotify has no "dislike" concept — nothing to click, not a failure to retry.
            return false;
        }

        var script = $$"""
            (() => {
                const el = {{elementExpression}};
                if (!el) return 'false';
                el.click();
                return 'true';
            })()
            """;
        var result = await RunScriptAsync(script).ConfigureAwait(true);
        return result == "true";
    }

    public async Task<NowPlayingInfo?> GetNowPlayingAsync()
    {
        var result = await RunScriptAsync(Profile.NowPlayingScript).ConfigureAwait(true);
        if (string.IsNullOrEmpty(result) || result == "null")
        {
            return null;
        }

        try
        {
            var parsed = JsonSerializer.Deserialize<NowPlayingJson>(result, JsonOptions);
            return parsed is null ? null : new NowPlayingInfo(parsed.Title, parsed.Artist, Profile.DisplayName);
        }
        catch (JsonException)
        {
            // Malformed/unexpected payload (e.g. the site's markup shifted mid-parse) — treat the
            // same as "nothing playing" rather than letting a bad chat command crash the app.
            return null;
        }
    }

    /// <summary>
    /// Runs a script on the UI thread and unwraps its result, then never lets a failure escape as
    /// an exception — a WebView call failing (page not loaded yet, the site's markup having
    /// shifted, a JS runtime error) should make one chat command report "didn't work", not take
    /// down the whole app (confirmed necessary via a real crash on !nowplaying, 2026-08-24, before
    /// this existed). Also guards against a real WebView2 quirk: ExecuteScriptAsync JSON-encodes
    /// its result even when the script already returns a string, so a script returning the literal
    /// string "true" can come back over the wire as the 6-character payload `"true"` (quotes
    /// included) rather than the 4-character `true`. Unwrapping once with Deserialize&lt;string&gt;
    /// handles that; a result that isn't itself a JSON string (shouldn't happen — every script
    /// here always returns a quoted string or null) falls back to the raw text unchanged.
    /// </summary>
    private async Task<string?> RunScriptAsync(string script)
    {
        try
        {
            var raw = await Dispatcher.UIThread.InvokeAsync(() => webView.InvokeScript(script)).ConfigureAwait(true);
            if (raw is null)
            {
                return null;
            }

            try
            {
                return JsonSerializer.Deserialize<string>(raw);
            }
            catch (JsonException)
            {
                return raw;
            }
        }
        catch (Exception)
        {
            return null;
        }
    }

    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    private sealed record NowPlayingJson(string Title, string Artist);
}
