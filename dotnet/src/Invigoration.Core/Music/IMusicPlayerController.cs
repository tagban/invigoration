namespace Invigoration.Core.Music;

/// <summary>
/// Bridges chat commands (<see cref="BotEngine"/>, UI-agnostic) to the music service Invigoration
/// controls — Spotify (<see cref="Spotify.SpotifyController"/>), with room for others such as
/// Plex. Every action returns false when it didn't happen (nothing playing, no active device, not
/// signed in), with <see cref="LastError"/> saying why, so callers can give a clear reply instead
/// of silently doing nothing.
/// </summary>
public interface IMusicPlayerController
{
    Task<bool> SkipAsync();

    /// <summary>Toggles play/pause — every service's own player exposes exactly one button for this, not separate play/pause/stop controls.</summary>
    Task<bool> PlayPauseAsync();

    Task<bool> ThumbsUpAsync();

    Task<bool> ThumbsDownAsync();

    /// <summary>What's loaded, playing or paused (see <see cref="NowPlayingInfo.IsPlaying"/>), or null when nothing is.</summary>
    Task<NowPlayingInfo?> GetNowPlayingAsync();

    /// <summary>
    /// Whether the service has a "like"/"dislike" concept at all — a service with no dislike
    /// should have !thumbsdown quietly do nothing rather than reply with a failure that implies a
    /// real, fixable problem.
    /// </summary>
    bool SupportsThumbsUp => true;

    bool SupportsThumbsDown => true;

    /// <summary>Why the most recent action didn't work, in words a chat reply can use — or null for "no particular reason to give".</summary>
    string? LastError => null;
}
