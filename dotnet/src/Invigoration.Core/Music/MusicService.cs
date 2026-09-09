namespace Invigoration.Core.Music;

/// <summary>
/// Which music service the embedded player tab is currently pointed at. Spotify and Pandora were
/// removed (2026-09-09, explicit request) — Spotify's web player was unreliably slow/hanging
/// ("spinning circle of death") in the embedded WebView, and Pandora was dropped alongside it to
/// keep the player to one well-supported service rather than three of uneven quality. Left as an
/// enum (not collapsed to a single constant) since MusicSettingsStore persists this by name/value
/// and IMusicPlayerController's shape doesn't assume a single implementation.
/// </summary>
public enum MusicService
{
    YouTubeMusic,
}
