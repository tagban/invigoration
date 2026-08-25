namespace Invigoration.Core.Music;

/// <summary>
/// Bridges chat commands (<see cref="BotEngine"/>, UI-agnostic) to the embedded YouTube Music
/// player window (Invigoration.App, the only place that can actually implement this — see
/// MusicPlayerRegistry). Every method returns false/null when the command didn't actually reach
/// a working player (e.g. the page hasn't finished loading, or the like/dislike click bounced to
/// a Google sign-in redirect because the user isn't signed in yet) so callers can give a clear
/// reply instead of silently doing nothing.
/// </summary>
public interface IMusicPlayerController
{
    Task<bool> SkipAsync();
    Task<bool> ThumbsUpAsync();
    Task<bool> ThumbsDownAsync();
    Task<NowPlayingInfo?> GetNowPlayingAsync();

    /// <summary>
    /// Whether the current service actually has a "like"/"dislike" concept at all — Spotify has
    /// no dislike (just a Save-to-Library heart), so !thumbsdown there should quietly do nothing
    /// rather than show a "couldn't dislike, make sure you're signed in" message that implies a
    /// real, fixable problem. Default true (every real implementation so far supports both) so
    /// existing/test controllers don't need updating just to add this.
    /// </summary>
    bool SupportsThumbsUp => true;

    bool SupportsThumbsDown => true;
}
