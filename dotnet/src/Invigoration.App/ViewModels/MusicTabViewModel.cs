using System.Collections.ObjectModel;
using System.Diagnostics;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.Models;
using Invigoration.Core.Music;
using Invigoration.Core.Music.Spotify;

namespace Invigoration.App.ViewModels;

/// <summary>
/// The Music tab: connects Invigoration to the user's Spotify and shows what's playing there, with
/// play/pause, skip and save. Spotify itself plays wherever the user listens (the Spotify app, a
/// phone, a speaker); this tab and every bot's music commands control it through the Web API
/// (SpotifyController, shared via MusicPlayerRegistry). The connection lives here rather than in
/// the view, so it's made once at startup from the saved sign-in and survives switching tabs.
/// Same duck-typed Title/HighlightBrush/etc. header shape as GlobalWhispersTabViewModel.
/// </summary>
public sealed partial class MusicTabViewModel : ViewModelBase
{
    public const string DeveloperDashboardUrl = "https://developer.spotify.com/dashboard";

    public const string WebPlayerUrl = "https://open.spotify.com";

    private static readonly TimeSpan SignInTimeout = TimeSpan.FromMinutes(5);

    /// <summary>How long Start waits for a Spotify app it opened to come online.</summary>
    private static readonly TimeSpan StartTimeout = TimeSpan.FromSeconds(30);

    private CancellationTokenSource? _connectCts;
    private string? _artworkUrl;

    public MusicTabViewModel()
    {
        ClientId = MusicSettingsStore.SpotifyClientId;
        if (ClientId.Length > 0 && MusicSettingsStore.SpotifyRefreshToken.Length > 0)
        {
            Use(new SpotifyController(ClientId, MusicSettingsStore.SpotifyRefreshToken, SaveRefreshToken));
        }
    }

    // --- Tab header ---

    public string Title => "";

    public IBrush HighlightBrush { get; } = new SolidColorBrush(Color.FromRgb(0x1E, 0xD7, 0x60));

    public Bitmap? TabIconImage => GameIconLoader.Get("spotify");

    public double HeaderFontSize => 13;

    public IBrush HeaderForeground => HighlightBrush;

    public bool HasUnread => false;

    // --- Connection ---

    public string RedirectUri => SpotifyAuthorization.RedirectUri;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(ConnectCommand))]
    public partial string ClientId { get; set; } = "";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ShowsSetup))]
    public partial bool IsConnected { get; set; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ShowsSetup))]
    [NotifyCanExecuteChangedFor(nameof(ConnectCommand))]
    public partial bool IsConnecting { get; set; }

    /// <summary>What's happening, or what went wrong — shown under the controls.</summary>
    [ObservableProperty]
    public partial string StatusText { get; set; } = "";

    public bool ShowsSetup => !IsConnected;

    // --- Now playing ---

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasNothingPlaying))]
    public partial bool HasTrack { get; set; }

    public bool HasNothingPlaying => !HasTrack;

    [ObservableProperty]
    public partial string TrackTitle { get; set; } = "";

    [ObservableProperty]
    public partial string Artist { get; set; } = "";

    [ObservableProperty]
    public partial string Album { get; set; } = "";

    [ObservableProperty]
    public partial string DeviceText { get; set; } = "";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(PlayPauseGlyph))]
    public partial bool IsPlaying { get; set; }

    public string PlayPauseGlyph => IsPlaying ? "⏸" : "▶";

    [ObservableProperty]
    public partial Bitmap? Artwork { get; set; }

    [ObservableProperty]
    public partial double Progress { get; set; }

    [ObservableProperty]
    public partial string TimeText { get; set; } = "";

    // --- Search ---

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SearchCommand))]
    public partial string SearchQuery { get; set; } = "";

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SearchCommand))]
    public partial bool IsSearching { get; set; }

    [ObservableProperty]
    public partial string SearchStatus { get; set; } = "";

    public ObservableCollection<MusicSearchResultViewModel> SearchResults { get; } = [];

    /// <summary>True while Start is opening Spotify and waiting for it.</summary>
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(StartSpotifyCommand))]
    public partial bool IsStarting { get; set; }

    private bool CanConnect() => !IsConnecting && ClientId.Trim().Length > 0;

    /// <summary>Opens Spotify's sign-in in the browser and waits for it to come back to this machine.</summary>
    [RelayCommand(CanExecute = nameof(CanConnect))]
    private async Task ConnectAsync()
    {
        var clientId = ClientId.Trim();
        var verifier = SpotifyAuthorization.CreateCodeVerifier();
        var state = SpotifyAuthorization.CreateState();
        _connectCts = new CancellationTokenSource(SignInTimeout);
        IsConnecting = true;
        StatusText = "Finish signing in to Spotify in your browser…";
        try
        {
            // Listening starts before the browser opens: an already-approved app comes straight back.
            var waitForCode = SpotifyAuthorization.WaitForCodeAsync(state, cancellationToken: _connectCts.Token);
            OpenInBrowser(SpotifyAuthorization.AuthorizeUri(clientId, SpotifyAuthorization.CodeChallenge(verifier), state).AbsoluteUri);
            var code = await waitForCode;

            StatusText = "Connecting…";
            var tokens = await SpotifyAuthorization.ExchangeCodeAsync(SpotifyController.SharedHttp, clientId, code, verifier, cancellationToken: _connectCts.Token);
            MusicSettingsStore.SpotifyClientId = clientId;
            MusicSettingsStore.SpotifyRefreshToken = tokens.RefreshToken;
            Use(new SpotifyController(clientId, tokens.RefreshToken, SaveRefreshToken, tokens: tokens));
            StatusText = "";
            await RefreshAsync();
        }
        catch (OperationCanceledException)
        {
            StatusText = _connectCts.IsCancellationRequested && !_cancelledByUser ? "Spotify sign-in timed out. Try again." : "";
        }
        catch (SpotifyAuthException ex)
        {
            StatusText = ex.Message;
        }
        finally
        {
            _cancelledByUser = false;
            IsConnecting = false;
            _connectCts.Dispose();
            _connectCts = null;
        }
    }

    private bool _cancelledByUser;

    [RelayCommand]
    private void CancelConnect()
    {
        _cancelledByUser = true;
        _connectCts?.Cancel();
    }

    /// <summary>Forgets the saved sign-in. Spotify keeps the app listed under the account's apps until removed there.</summary>
    [RelayCommand]
    private void Disconnect()
    {
        MusicSettingsStore.SpotifyRefreshToken = "";
        MusicPlayerRegistry.Controller = null;
        IsConnected = false;
        ClearTrack();
        SearchResults.Clear();
        SearchStatus = "";
        StatusText = "Disconnected from Spotify.";
    }

    [RelayCommand]
    private static void OpenDeveloperDashboard() => OpenInBrowser(DeveloperDashboardUrl);

    [RelayCommand]
    private Task PlayPause() => ActAsync(c => c.PlayPauseAsync(), startSpotifyIfClosed: true);

    [RelayCommand]
    private Task Skip() => ActAsync(c => c.SkipAsync());

    [RelayCommand]
    private Task Save() => ActAsync(c => c.ThumbsUpAsync(), "Saved to your Spotify library.");

    private bool CanSearch() => !IsSearching && SearchQuery.Trim().Length > 0;

    [RelayCommand(CanExecute = nameof(CanSearch))]
    private async Task SearchAsync()
    {
        if (MusicPlayerRegistry.Controller is not SpotifyController controller)
        {
            return;
        }

        IsSearching = true;
        SearchStatus = "Searching…";
        try
        {
            var tracks = await controller.SearchTracksAsync(SearchQuery);
            SearchResults.Clear();
            if (tracks is null)
            {
                SearchStatus = controller.LastError ?? "Search didn't work.";
                return;
            }

            foreach (var track in tracks)
            {
                var result = new MusicSearchResultViewModel(track, PlayTrackAsync, QueueTrackAsync);
                SearchResults.Add(result);
                _ = result.LoadArtworkAsync();
            }

            SearchStatus = tracks.Count == 0 ? "No tracks found." : "";
        }
        finally
        {
            IsSearching = false;
        }
    }

    private Task PlayTrackAsync(SpotifyTrack track) =>
        ActAsync(c => ((SpotifyController)c).PlayTrackAsync(track.Uri), $"Playing {track.Title}.", startSpotifyIfClosed: true);

    private Task QueueTrackAsync(SpotifyTrack track) =>
        ActAsync(c => ((SpotifyController)c).QueueTrackAsync(track.Uri), $"Added {track.Title} to the queue.", startSpotifyIfClosed: true);

    private bool CanStartSpotify() => !IsStarting;

    /// <summary>Opens Spotify (see OpenSpotify), waits for it to come online, and starts playing where it left off.</summary>
    [RelayCommand(CanExecute = nameof(CanStartSpotify))]
    private Task StartSpotify() => ActAsync(c => ((SpotifyController)c).ResumeAsync(), startSpotifyIfClosed: true);

    /// <summary>Reads what's playing — polled by the Music tab while it's showing.</summary>
    public async Task RefreshAsync()
    {
        if (MusicPlayerRegistry.Controller is not SpotifyController controller)
        {
            return;
        }

        var nowPlaying = await controller.GetNowPlayingAsync();
        if (controller.SignInExpired)
        {
            IsConnected = false;
            ClearTrack();
            StatusText = "Spotify sign-in has expired — connect again.";
            return;
        }

        if (nowPlaying is null)
        {
            ClearTrack();
            if (controller.LastError is { } error)
            {
                StatusText = error;
            }

            return;
        }

        HasTrack = true;
        TrackTitle = nowPlaying.Title;
        Artist = nowPlaying.Artist;
        Album = nowPlaying.Album;
        IsPlaying = nowPlaying.IsPlaying;
        DeviceText = nowPlaying.Device is { Length: > 0 } device ? $"{(nowPlaying.IsPlaying ? "Playing" : "Paused")} on {device}" : "";
        Progress = nowPlaying.DurationMs > 0 ? Math.Clamp((double)nowPlaying.ProgressMs / nowPlaying.DurationMs, 0, 1) : 0;
        TimeText = nowPlaying.DurationMs > 0 ? $"{FormatTime(nowPlaying.ProgressMs)} / {FormatTime(nowPlaying.DurationMs)}" : "";
        await LoadArtworkAsync(nowPlaying.ArtworkUrl);
    }

    private void Use(SpotifyController controller)
    {
        MusicPlayerRegistry.Controller = controller;
        IsConnected = true;
    }

    /// <param name="startSpotifyIfClosed">When the action fails because no Spotify app is open, open one, wait for it, and try again.</param>
    private async Task ActAsync(Func<IMusicPlayerController, Task<bool>> action, string? successText = null, bool startSpotifyIfClosed = false)
    {
        if (MusicPlayerRegistry.Controller is not { } controller)
        {
            return;
        }

        var ok = await action(controller);
        if (!ok && startSpotifyIfClosed && controller is SpotifyController { NoDeviceOpen: true } spotify && !IsStarting)
        {
            IsStarting = true;
            try
            {
                StatusText = "Opening Spotify…";
                OpenSpotify();
                var deadline = DateTimeOffset.UtcNow + StartTimeout;
                while (DateTimeOffset.UtcNow < deadline && !await spotify.HasOpenDeviceAsync())
                {
                    await Task.Delay(TimeSpan.FromSeconds(2));
                }

                ok = await action(controller);
                if (!ok && spotify.NoDeviceOpen)
                {
                    StatusText = "Spotify didn't come online. If it opened in your browser, sign in there, then press play again.";
                    await RefreshAsync();
                    return;
                }
            }
            finally
            {
                IsStarting = false;
            }
        }

        StatusText = ok ? successText ?? "" : controller.LastError ?? "That didn't work.";
        await RefreshAsync();
    }

    /// <summary>
    /// Opens something Spotify can play on: on a Mac, the Spotify app if it's installed — hidden and
    /// in the background, so the music just starts — otherwise the web player in the default browser.
    /// Elsewhere, the Spotify app via its spotify: link, which the OS hands to the web player's site
    /// if the app isn't there.
    /// </summary>
    private static void OpenSpotify()
    {
        try
        {
            if (OperatingSystem.IsMacOS())
            {
                var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                if (Directory.Exists("/Applications/Spotify.app") || Directory.Exists(Path.Combine(home, "Applications", "Spotify.app")))
                {
                    Process.Start("open", ["-g", "-j", "-a", "Spotify"]);
                    return;
                }

                OpenInBrowser(WebPlayerUrl);
                return;
            }

            OpenInBrowser("spotify:");
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException)
        {
            OpenInBrowser(WebPlayerUrl);
        }
    }

    private void ClearTrack()
    {
        HasTrack = false;
        TrackTitle = Artist = Album = DeviceText = TimeText = "";
        IsPlaying = false;
        Progress = 0;
        Artwork = null;
        _artworkUrl = null;
    }

    private async Task LoadArtworkAsync(string? url)
    {
        if (url == _artworkUrl)
        {
            return;
        }

        _artworkUrl = url;
        if (url is null)
        {
            Artwork = null;
            return;
        }

        try
        {
            var bytes = await SpotifyController.SharedHttp.GetByteArrayAsync(url);
            if (_artworkUrl == url)
            {
                Artwork = new Bitmap(new MemoryStream(bytes));
            }
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or ArgumentException)
        {
            Artwork = null;
        }
    }

    private static void SaveRefreshToken(string refreshToken) => MusicSettingsStore.SpotifyRefreshToken = refreshToken;

    private static string FormatTime(long ms) => TimeSpan.FromMilliseconds(ms) is var t && t.TotalHours >= 1
        ? t.ToString(@"h\:mm\:ss")
        : t.ToString(@"m\:ss");

    private static void OpenInBrowser(string url) => Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
}
