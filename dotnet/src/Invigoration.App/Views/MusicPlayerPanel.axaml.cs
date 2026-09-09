using Avalonia.Controls;
using Avalonia.Platform;
using Invigoration.App.Models;
using Invigoration.App.Music;
using Invigoration.App.ViewModels;
using Invigoration.Core.Config;
using Invigoration.Core.Music;

namespace Invigoration.App.Views;

/// <summary>
/// The embedded music player — a NativeWebView pointed at YouTube Music, driven by chat commands
/// via WebViewMusicController (registered as the process-wide MusicPlayerRegistry.Controller).
/// Deliberately NOT the Music tab's actual TabControl content (see MusicTabView, its trivial
/// placeholder) — Avalonia's TabControl destroys/detaches a non-selected tab's content by default,
/// which killed playback the moment you switched to a different tab (confirmed live, 2026-08-24).
/// Instead this lives as a permanent sibling overlay in MainWindow.axaml, positioned over the
/// TabControl and shown/hidden purely via IsVisible (MainWindowViewModel.IsMusicTabSelected) — the
/// control, and the underlying native WebView2 handle, is created once and never destroyed for the
/// app's lifetime, so playback keeps going no matter which tab is actually showing.
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
        _viewModel = DataContext as MusicTabViewModel;
        if (_viewModel is null)
        {
            return;
        }

        _controller ??= new WebViewMusicController(WebView);
        MusicPlayerRegistry.Controller = _controller;

        var profile = MusicServiceProfile.YouTubeMusic;
        _controller.Profile = profile;
        WebView.UserAgent = profile.MobileUserAgent ?? "";
        WebView.Source = new Uri(profile.HomeUrl);
    }
}
