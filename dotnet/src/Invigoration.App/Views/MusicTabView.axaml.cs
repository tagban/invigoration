using Avalonia;
using Avalonia.Controls;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Invigoration.App.ViewModels;

namespace Invigoration.App.Views;

/// <summary>The Music tab (see MusicTabViewModel). Polls Spotify for what's playing every few seconds, only while the tab is on screen.</summary>
public partial class MusicTabView : UserControl
{
    private static readonly TimeSpan RefreshInterval = TimeSpan.FromSeconds(3);

    private readonly DispatcherTimer _refreshTimer = new() { Interval = RefreshInterval };

    public MusicTabView()
    {
        InitializeComponent();
        _refreshTimer.Tick += async (_, _) => await RefreshAsync();
    }

    protected override async void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnAttachedToVisualTree(e);
        _refreshTimer.Start();
        await RefreshAsync();
    }

    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnDetachedFromVisualTree(e);
        _refreshTimer.Stop();
    }

    private async Task RefreshAsync()
    {
        if (DataContext is MusicTabViewModel { IsConnected: true } vm)
        {
            await vm.RefreshAsync();
        }
    }

    private async void OnCopyRedirectUriClick(object? sender, RoutedEventArgs e)
    {
        if (TopLevel.GetTopLevel(this)?.Clipboard is { } clipboard && DataContext is MusicTabViewModel vm)
        {
            await clipboard.SetTextAsync(vm.RedirectUri);
        }
    }
}
