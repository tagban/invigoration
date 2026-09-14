namespace Invigoration.Core.Music;

/// <summary>
/// What's loaded in the music player right now, as reported by whatever implements
/// <see cref="IMusicPlayerController"/>. Service is a display name (e.g. "Spotify"). The rest is
/// optional detail for the Music tab and player bar; chat only needs Title and Artist.
/// <see cref="IsPlaying"/> is false for a track that's loaded but paused.
/// </summary>
public sealed record NowPlayingInfo(string Title, string Artist, string? Service = null)
{
    public string Album { get; init; } = "";

    public string? ArtworkUrl { get; init; }

    public bool IsPlaying { get; init; } = true;

    /// <summary>Which of the user's devices is playing it ("MacBook Pro", "Kitchen speaker").</summary>
    public string? Device { get; init; }

    public long ProgressMs { get; init; }

    public long DurationMs { get; init; }

    /// <summary>The service's own id for the track (a spotify: URI), used to save it.</summary>
    public string? Uri { get; init; }
}
