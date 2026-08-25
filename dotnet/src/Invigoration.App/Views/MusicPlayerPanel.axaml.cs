using Avalonia.Controls;
using Avalonia.Platform;
using Avalonia.Interactivity;
using Invigoration.App.Models;
using Invigoration.App.Music;
using Invigoration.App.ViewModels;
using Invigoration.Core.Config;
using Invigoration.Core.Music;

namespace Invigoration.App.Views;

/// <summary>
/// The embedded music player — a NativeWebView pointed at whichever service is selected, driven
/// by chat commands via WebViewMusicController (registered as the process-wide
/// MusicPlayerRegistry.Controller). Deliberately NOT the Music tab's actual TabControl content
/// (see MusicTabView, its trivial placeholder) — Avalonia's TabControl destroys/detaches a
/// non-selected tab's content by default, which killed playback the moment you switched to a
/// different tab (confirmed live, 2026-08-24). Instead this lives as a permanent sibling overlay
/// in MainWindow.axaml, positioned over the TabControl and shown/hidden purely via IsVisible
/// (MainWindowViewModel.IsMusicTabSelected) — the control, and the underlying native WebView2
/// handle, is created once and never destroyed for the app's lifetime, so playback keeps going
/// no matter which tab is actually showing.
/// </summary>
public partial class MusicPlayerPanel : UserControl
{
    private WebViewMusicController? _controller;
    private MusicTabViewModel? _viewModel;

    public MusicPlayerPanel()
    {
        InitializeComponent();
        DataContextChanged += (_, _) => Attach();
        WebView.EnvironmentRequested += OnEnvironmentRequested;
        WebView.NavigationCompleted += OnNavigationCompleted;
    }

    /// <summary>
    /// Points WebView2's profile at our own AppData folder instead of whatever default location
    /// it'd otherwise pick — without this, a self-contained single-file published exe has no
    /// guarantee its default profile path stays stable across runs, which is exactly what made
    /// login not survive an app restart in the first (popup-window) version of this feature.
    /// Windows-only (WebView2); macOS/Linux backends persist via their own OS-level webview
    /// storage without needing this.
    /// </summary>
    private static void OnEnvironmentRequested(object? sender, WebViewEnvironmentRequestedEventArgs e)
    {
        if (e is WindowsWebView2EnvironmentRequestedEventArgs webView2)
        {
            webView2.UserDataFolder = Path.Combine(ConfigStore.DefaultConfigDirectory(), "MusicPlayerProfile");
        }
    }

    /// <summary>
    /// Best-effort, not a guaranteed bandwidth-saving audio-only stream — YouTube Music's web
    /// player doesn't expose a real "audio only" API/toggle, so this just visually hides the
    /// video element after each navigation so the tab always shows album art instead of playing
    /// video, which is what was actually asked for.
    /// </summary>
    private void OnNavigationCompleted(object? sender, WebViewNavigationCompletedEventArgs e)
    {
        if (_viewModel?.SelectedService != MusicService.YouTubeMusic)
        {
            return;
        }

        const string css = "video { visibility: hidden !important; }";
        _ = WebView.InvokeScript($$"""
            (() => {
                const style = document.createElement('style');
                style.textContent = {{System.Text.Json.JsonSerializer.Serialize(css)}};
                document.head.appendChild(style);
                return 'true';
            })()
            """);
    }

    private void Attach()
    {
        if (_viewModel is not null)
        {
            _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
        }

        _viewModel = DataContext as MusicTabViewModel;
        if (_viewModel is null)
        {
            return;
        }

        _viewModel.PropertyChanged += OnViewModelPropertyChanged;
        _controller ??= new WebViewMusicController(WebView);
        MusicPlayerRegistry.Controller = _controller;
        ApplyService(_viewModel.SelectedService);
    }

    private void OnViewModelPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(MusicTabViewModel.SelectedService) && _viewModel is not null)
        {
            ApplyService(_viewModel.SelectedService);
        }
    }

    private void ApplyService(MusicService service)
    {
        var profile = MusicServiceProfile.For(service);
        if (_controller is not null)
        {
            _controller.Profile = profile;
        }

        // Must be set before Source so it's in effect for the very first navigation request — see
        // MusicServiceProfile.Spotify's remarks for why this (and only Spotify, so far) needs it.
        // "" (not a hardcoded desktop UA string) restores WebView2's own real default — confirmed
        // live this was the actual cause of a real regression: a hardcoded "Chrome/124..." string
        // applied to every service (not just Spotify) made YouTube Music stop reporting
        // now-playing entirely, almost certainly because that fabricated version string stopped
        // matching WebView2's real Sec-CH-UA client-hints headers, which YouTube checks. Per
        // WebView2's own documented behavior, an empty UserAgent means "use the default" — not a
        // blank header — so this is the correct way to un-spoof, not a fallback guess.
        WebView.UserAgent = profile.MobileUserAgent ?? "";
        WebView.Source = new Uri(profile.HomeUrl);
        YouTubeMusicIcon.Source = GameIconLoader.Get("youtube-music");
        SpotifyIcon.Source = GameIconLoader.Get("spotify");
        PandoraIcon.Source = GameIconLoader.Get("pandora");
        YouTubeMusicButton.Opacity = service == MusicService.YouTubeMusic ? 1.0 : 0.4;
        SpotifyButton.Opacity = service == MusicService.Spotify ? 1.0 : 0.4;
        PandoraButton.Opacity = service == MusicService.Pandora ? 1.0 : 0.4;

        if (service == MusicService.Spotify)
        {
            _ = ReapplySpotifyOnceSettledAsync();
        }
    }

    /// <summary>
    /// Confirmed live: Spotify's own mobile-layout detection sometimes misses on the very first
    /// navigation into it (still shows the desktop layout squeezed into the panel) but is fine
    /// right after — switching to Pandora and back to Spotify "looks great". That first
    /// navigation can race either the WebView2 environment still spinning up or this panel's own
    /// layout/sizing not having settled yet (it's an always-alive overlay shown/hidden via
    /// IsVisible, not freshly created). Re-navigating once more, shortly after, reproduces the
    /// same fix automatically instead of requiring the user to manually switch tabs and back.
    /// </summary>
    private async Task ReapplySpotifyOnceSettledAsync()
    {
        await Task.Delay(TimeSpan.FromSeconds(1)).ConfigureAwait(true);
        if (_viewModel?.SelectedService == MusicService.Spotify)
        {
            WebView.UserAgent = MusicServiceProfile.Spotify.MobileUserAgent ?? "";
            WebView.Source = new Uri(MusicServiceProfile.Spotify.HomeUrl);
        }
    }

    private void OnYouTubeMusicClick(object? sender, RoutedEventArgs e)
    {
        if (_viewModel is not null)
        {
            _viewModel.SelectedService = MusicService.YouTubeMusic;
        }
    }

    private void OnSpotifyClick(object? sender, RoutedEventArgs e)
    {
        if (_viewModel is not null)
        {
            _viewModel.SelectedService = MusicService.Spotify;
        }
    }

    private void OnPandoraClick(object? sender, RoutedEventArgs e)
    {
        if (_viewModel is not null)
        {
            _viewModel.SelectedService = MusicService.Pandora;
        }
    }
}
