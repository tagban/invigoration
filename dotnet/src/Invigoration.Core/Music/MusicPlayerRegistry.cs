namespace Invigoration.Core.Music;

/// <summary>
/// The one music controller the whole app shares — every bot's (and Hotline session's) !skip/
/// !thumbsup/!nowplaying controls the same Spotify account, whichever tab the command came from,
/// and a reply only needs to go back where the command came from. Set by the App's Music tab when
/// Spotify is connected (Core stays UI-agnostic) and null otherwise, so a command sent while
/// nothing's connected gets <see cref="NotConnectedReply"/> instead of a NullReferenceException.
/// </summary>
public static class MusicPlayerRegistry
{
    public const string NotConnectedReply = "Spotify isn't connected — connect it in Invigoration's Music tab.";

    public static IMusicPlayerController? Controller { get; set; }

    /// <summary>The chat reply for !nowplaying — shared by bots and Hotline so they say the same thing.</summary>
    public static async Task<string> DescribeNowPlayingAsync()
    {
        if (Controller is not { } controller)
        {
            return NotConnectedReply;
        }

        var nowPlaying = await controller.GetNowPlayingAsync().ConfigureAwait(false);
        if (nowPlaying is null)
        {
            return controller.LastError ?? "Nothing seems to be playing.";
        }

        var service = string.IsNullOrEmpty(nowPlaying.Service) ? "" : $" on {nowPlaying.Service}";
        return nowPlaying.IsPlaying
            ? $"/me is now playing {nowPlaying.Title} - by {nowPlaying.Artist}{service}."
            : $"/me has {nowPlaying.Title} - by {nowPlaying.Artist} paused{service}.";
    }
}
