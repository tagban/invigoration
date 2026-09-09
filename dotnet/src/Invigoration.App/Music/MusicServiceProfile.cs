using Invigoration.Core.Music;

namespace Invigoration.App.Music;

/// <summary>
/// Per-service DOM-scripting knowledge for WebViewMusicController — raw JS snippets rather than
/// individual CSS-selector strings, since a different service's now-playing markup could be
/// shaped too differently to squeeze into one generic template. Only YouTube Music remains as of
/// 2026-09-09 (Spotify and Pandora removed — see MusicService's remarks), but the shape stays
/// per-service rather than collapsing to one hardcoded set of scripts. LikeScript/DislikeScript
/// are nullable for a service with no "dislike" concept in its web player at all.
/// </summary>
public sealed record MusicServiceProfile(
    MusicService Service,
    string DisplayName,
    string IconKey,
    string HomeUrl,
    string NextScript,
    string PlayPauseScript,
    string? LikeScript,
    string? DislikeScript,
    string NowPlayingScript,
    string? MobileUserAgent = null)
{
    /// <summary>
    /// Confirmed live against the real site while building this (2026-08-24, via a browser tool —
    /// not guessed): player bar is ytmusic-player-bar; .title/.byline for now-playing (byline is
    /// "Artist • N views • N likes", split on " • " and take the first segment); .next-button;
    /// like/dislike are #like-button-renderer button[aria-label="Like"/"Dislike"].
    /// </summary>
    public static readonly MusicServiceProfile YouTubeMusic = new(
        MusicService.YouTubeMusic,
        "YouTube Music",
        "youtube-music",
        "https://music.youtube.com",
        NextScript: """document.querySelector('ytmusic-player-bar .next-button')""",
        // Not separately confirmed live like the other selectors on this profile (added
        // 2026-08-26 for the pause/play command) — #play-pause-button is YouTube Music's
        // well-known, stable player-bar element id, same one youtube-music-desktop-app and
        // similar community wrappers already rely on.
        PlayPauseScript: """document.querySelector('ytmusic-player-bar #play-pause-button')""",
        LikeScript: """document.querySelector('ytmusic-player-bar #like-button-renderer button[aria-label="Like"]')""",
        DislikeScript: """document.querySelector('ytmusic-player-bar #like-button-renderer button[aria-label="Dislike"]')""",
        NowPlayingScript: """
            (() => {
                const bar = document.querySelector('ytmusic-player-bar');
                const title = bar?.querySelector('.title')?.textContent?.trim() ?? '';
                if (!title) return 'null';
                const bylineRaw = bar.querySelector('.byline')?.textContent?.trim() ?? '';
                const artist = bylineRaw.split(' • ')[0] ?? '';
                return JSON.stringify({ title, artist });
            })()
            """);

    public static MusicServiceProfile For(MusicService service) => service switch
    {
        _ => YouTubeMusic,
    };
}
