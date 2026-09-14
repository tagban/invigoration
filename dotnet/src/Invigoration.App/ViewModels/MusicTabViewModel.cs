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

    private static readonly TimeSpan SignInTimeout = TimeSpan.FromMinutes(5);

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
        StatusText = "Disconnected from Spotify.";
    }

    [RelayCommand]
    private static void OpenDeveloperDashboard() => OpenInBrowser(DeveloperDashboardUrl);

    [RelayCommand]
    private Task PlayPause() => ActAsync(c => c.PlayPauseAsync());

    [RelayCommand]
    private Task Skip() => ActAsync(c => c.SkipAsync());

    [RelayCommand]
    private Task Save() => ActAsync(c => c.ThumbsUpAsync(), "Saved to your Spotify library.");

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

    private async Task ActAsync(Func<IMusicPlayerController, Task<bool>> action, string? successText = null)
    {
        if (MusicPlayerRegistry.Controller is not { } controller)
        {
            return;
        }

        var ok = await action(controller);
        StatusText = ok ? successText ?? "" : controller.LastError ?? "That didn't work.";
        await RefreshAsync();
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
