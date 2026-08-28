using Invigoration.Core.Music;

namespace Invigoration.App.Music;

/// <summary>
/// Per-service DOM-scripting knowledge for WebViewMusicController — raw JS snippets rather than
/// individual CSS-selector strings, since YouTube Music's and Spotify's now-playing markup are
/// shaped too differently to squeeze into one generic template. LikeScript/DislikeScript are
/// nullable: Spotify has no "dislike" concept in its web player at all (just a Save-to-Library
/// heart), so that command just replies "not supported" there instead of pretending an
/// equivalent exists.
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

    /// <summary>
    /// Confirmed working live this session — selectors below are real. MobileUserAgent is the fix
    /// for a real, confirmed-live difference from YouTube Music/Pandora: those two lay out compact
    /// purely from CSS width (any desktop UA squeezed into ~420px gets the mobile layout), but
    /// open.spotify.com only serves its actual mobile/app-style layout to a UA it detects as a real
    /// mobile device — WebView2's default desktop UA at the same width instead renders the full
    /// desktop chrome squeezed into a narrow column. Spoofing a real Android Chrome UA (matching
    /// what a real device would send) gets Spotify to serve the compact layout the same way the
    /// other two already do natively.
    /// </summary>
    public static readonly MusicServiceProfile Spotify = new(
        MusicService.Spotify,
        "Spotify",
        "spotify",
        "https://open.spotify.com",
        NextScript: """document.querySelector('[data-testid="control-button-skip-forward"]')""",
        // "control-button-playpause" is one of Spotify's long-stable data-testids (same family
        // as skip-forward above, widely relied on by community tools like spicetify) — not
        // separately live-confirmed here, but a much safer bet than most of Spotify's markup.
        PlayPauseScript: """document.querySelector('[data-testid="control-button-playpause"]')""",
        LikeScript: """document.querySelector('[data-testid="add-button"]')""",
        DislikeScript: null,
        NowPlayingScript: """
            (() => {
                const title = document.querySelector('[data-testid="context-item-info-title"]')?.textContent?.trim() ?? '';
                if (!title) return 'null';
                const artist = document.querySelector('[data-testid="context-item-info-subtitles"]')?.textContent?.trim() ?? '';
                return JSON.stringify({ title, artist });
            })()
            """,
        MobileUserAgent: "Mozilla/5.0 (Linux; Android 14; Pixel 8) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Mobile Safari/537.36");

    /// <summary>
    /// Confirmed live against a real signed-in station (not guessed): now-playing title/artist
    /// are .Tuner__Audio__TrackDetail__title/__artist; skip/thumbs are aria-labeled "Skip
    /// forwards"/"Thumb Up this song"/"Thumb Down this song". Thumbs only appear while playing an
    /// actual station (radio) — Pandora's on-demand album/playlist playback has no rating concept,
    /// same shape as Spotify having no dislike; that's a real Pandora product distinction; not
    /// something to special-case here since ThumbsUp/DownAsync already just report failure/absence
    /// naturally when the buttons aren't in the DOM.
    /// </summary>
    public static readonly MusicServiceProfile Pandora = new(
        MusicService.Pandora,
        "Pandora",
        "pandora",
        "https://www.pandora.com",
        NextScript: """document.querySelector('[aria-label="Skip forwards"]')""",
        // Not separately confirmed live (added 2026-08-26) — Pandora's play/pause button's
        // aria-label flips between "Pause" and "Play" depending on state, matching the same
        // aria-label convention as the confirmed Skip/Thumb buttons above.
        PlayPauseScript: """document.querySelector('[aria-label="Pause"]') ?? document.querySelector('[aria-label="Play"]')""",
        LikeScript: """document.querySelector('[aria-label="Thumb Up this song"]')""",
        DislikeScript: """document.querySelector('[aria-label="Thumb Down this song"]')""",
        NowPlayingScript: """
            (() => {
                const title = document.querySelector('.Tuner__Audio__TrackDetail__title')?.textContent?.trim() ?? '';
                if (!title) return 'null';
                const artist = document.querySelector('.Tuner__Audio__TrackDetail__artist')?.textContent?.trim() ?? '';
                return JSON.stringify({ title, artist });
            })()
            """);

    public static MusicServiceProfile For(MusicService service) => service switch
    {
        MusicService.Spotify => Spotify,
        MusicService.Pandora => Pandora,
        _ => YouTubeMusic,
    };
}
