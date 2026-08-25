using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.Models;
using Invigoration.Core.Music;

namespace Invigoration.App.ViewModels;

/// <summary>
/// Backs the optional bottom playback-control bar (MusicBarView) — a thin strip docked below the
/// whole window's content, visible no matter which top-level tab is showing, so playback can be
/// controlled without switching to the Music tab itself (explicit request, 2026-08-24: "it would
/// make it easier to control if I'm looking at the bot"). Purely a read/act-through layer over
/// MusicPlayerRegistry.Controller — refreshed by MainWindow.axaml.cs's existing title-bar polling
/// timer (RefreshAsync), not its own separate poll loop.
/// </summary>
public sealed partial class MusicBarViewModel : ViewModelBase
{
    [ObservableProperty]
    public partial string Title { get; set; } = "";

    [ObservableProperty]
    public partial string Artist { get; set; } = "";

    [ObservableProperty]
    public partial bool HasNowPlaying { get; set; }

    [ObservableProperty]
    public partial Bitmap? ServiceIcon { get; set; }

    [ObservableProperty]
    public partial bool SupportsThumbsUp { get; set; } = true;

    [ObservableProperty]
    public partial bool SupportsThumbsDown { get; set; } = true;

    public async Task RefreshAsync()
    {
        var controller = MusicPlayerRegistry.Controller;
        if (controller is null)
        {
            HasNowPlaying = false;
            Title = "";
            Artist = "";
            return;
        }

        SupportsThumbsUp = controller.SupportsThumbsUp;
        SupportsThumbsDown = controller.SupportsThumbsDown;
        ServiceIcon = GameIconLoader.Get(MusicSettingsStore.SelectedService switch
        {
            MusicService.Spotify => "spotify",
            MusicService.Pandora => "pandora",
            _ => "youtube-music",
        });

        var nowPlaying = await controller.GetNowPlayingAsync().ConfigureAwait(true);
        HasNowPlaying = nowPlaying is not null;
        Title = nowPlaying?.Title ?? "";
        Artist = nowPlaying?.Artist ?? "";
    }

    [RelayCommand]
    private Task Skip() => MusicPlayerRegistry.Controller?.SkipAsync() ?? Task.FromResult(false);

    [RelayCommand]
    private Task ThumbsUp() => MusicPlayerRegistry.Controller?.ThumbsUpAsync() ?? Task.FromResult(false);

    [RelayCommand]
    private Task ThumbsDown() => MusicPlayerRegistry.Controller?.ThumbsDownAsync() ?? Task.FromResult(false);
}
