namespace Invigoration.Core.Music.Spotify;

/// <summary>A track found by <see cref="SpotifyController.SearchTracksAsync"/>. ArtworkUrl is a thumbnail-sized album image.</summary>
public sealed record SpotifyTrack(string Uri, string Title, string Artist, string Album, string? ArtworkUrl, long DurationMs);
