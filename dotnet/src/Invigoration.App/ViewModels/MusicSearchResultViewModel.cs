using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.Core.Music.Spotify;

namespace Invigoration.App.ViewModels;

/// <summary>One track in the Music tab's search results, with its own Play and Queue buttons.</summary>
public sealed partial class MusicSearchResultViewModel(SpotifyTrack track, Func<SpotifyTrack, Task> play, Func<SpotifyTrack, Task> queue) : ViewModelBase
{
    public SpotifyTrack Track { get; } = track;

    public string Title => Track.Title;

    public string Artist => Track.Artist;

    public string DurationText => Track.DurationMs > 0 ? TimeSpan.FromMilliseconds(Track.DurationMs).ToString(@"m\:ss") : "";

    [ObservableProperty]
    public partial Bitmap? Artwork { get; set; }

    [RelayCommand]
    private Task Play() => play(Track);

    [RelayCommand]
    private Task Queue() => queue(Track);

    public async Task LoadArtworkAsync()
    {
        if (Track.ArtworkUrl is not { } url)
        {
            return;
        }

        try
        {
            Artwork = new Bitmap(new MemoryStream(await SpotifyController.SharedHttp.GetByteArrayAsync(url)));
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or ArgumentException)
        {
            Artwork = null;
        }
    }
}
