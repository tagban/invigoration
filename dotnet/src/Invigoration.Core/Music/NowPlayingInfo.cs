namespace Invigoration.Core.Music;

/// <summary>The currently-playing track, as reported by whatever's implementing <see cref="IMusicPlayerController"/>. Service is a display name (e.g. "Spotify", "YouTube Music"), not the MusicService enum, since Core doesn't need to know the App-layer service list.</summary>
public sealed record NowPlayingInfo(string Title, string Artist, string? Service = null);
