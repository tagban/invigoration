using System.Collections.ObjectModel;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.Models;
using Invigoration.App.Music;
using Invigoration.Core.Music;
using Invigoration.Core.Music.Pandora;

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

    /// <summary>True only while the registered controller is actually a PandoraPlayerController — Pandora is a real API now (not a WebView the user already has full station-picking UI inside), so the bottom bar can drive station selection directly instead of requiring a trip to the Music tab. See MusicPlayerPanel's own station picker, which this mirrors.</summary>
    [ObservableProperty]
    public partial bool IsPandora { get; set; }

    public ObservableCollection<PandoraStation> Stations { get; } = [];

    [ObservableProperty]
    public partial PandoraStation? SelectedStation { get; set; }

    private PandoraPlayerController? _pandoraController;
    private bool _suppressStationChange;

    public async Task RefreshAsync()
    {
        var controller = MusicPlayerRegistry.Controller;
        if (controller is null)
        {
            HasNowPlaying = false;
            Title = "";
            Artist = "";
            IsPandora = false;
            _pandoraController = null;
            Stations.Clear();
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

        IsPandora = controller is PandoraPlayerController;
        if (controller is PandoraPlayerController pandora)
        {
            await SyncPandoraStationsAsync(pandora).ConfigureAwait(true);
        }
        else
        {
            _pandoraController = null;
            Stations.Clear();
        }
    }

    /// <summary>Fetches the station list once per logged-in controller instance (not on every poll tick — a real network call) and keeps SelectedStation in sync with whatever's actually playing, including a station started from the Music tab's own picker.</summary>
    private async Task SyncPandoraStationsAsync(PandoraPlayerController pandora)
    {
        if (!pandora.IsLoggedIn)
        {
            _pandoraController = null;
            Stations.Clear();
            return;
        }

        if (!ReferenceEquals(_pandoraController, pandora))
        {
            _pandoraController = pandora;
            Stations.Clear();
            foreach (var station in await pandora.GetStationsAsync().ConfigureAwait(true))
            {
                Stations.Add(station);
            }
        }

        var current = Stations.FirstOrDefault(s => s.StationToken == pandora.CurrentStationToken);
        if (!ReferenceEquals(SelectedStation, current))
        {
            _suppressStationChange = true;
            SelectedStation = current;
            _suppressStationChange = false;
        }
    }

    partial void OnSelectedStationChanged(PandoraStation? value)
    {
        if (_suppressStationChange || value is null || _pandoraController is null || value.StationToken == _pandoraController.CurrentStationToken)
        {
            return;
        }

        _ = _pandoraController.PlayStationAsync(value.StationToken);
    }

    [RelayCommand]
    private Task Skip() => MusicPlayerRegistry.Controller?.SkipAsync() ?? Task.FromResult(false);

    /// <summary>Toggles play/pause — added 2026-08-26 per explicit request ("Need to add Stop/Pause/Play Controls on the bottom player bar"). One button, not three: every service's own player exposes exactly one play/pause toggle, not a separate stop control (see IMusicPlayerController.PlayPauseAsync's remarks).</summary>
    [RelayCommand]
    private Task PlayPause() => MusicPlayerRegistry.Controller?.PlayPauseAsync() ?? Task.FromResult(false);

    [RelayCommand]
    private Task ThumbsUp() => MusicPlayerRegistry.Controller?.ThumbsUpAsync() ?? Task.FromResult(false);

    [RelayCommand]
    private Task ThumbsDown() => MusicPlayerRegistry.Controller?.ThumbsDownAsync() ?? Task.FromResult(false);
}
