using Avalonia.Media;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.App.Models;
using Invigoration.Core.Music;

namespace Invigoration.App.ViewModels;

/// <summary>
/// A persistent top-level tab (see MainWindowViewModel.TopLevelTabs) hosting the embedded music
/// player — same "duck-typed Title/HighlightBrush/etc." pattern as GlobalWhispersTabViewModel, so
/// it renders in the same TabControl header template without needing a shared base type. Always
/// present (added once in RefreshTopLevelTabs, same as the Whispers tab), unlike bot tabs which
/// come and go with Bots.
/// </summary>
public sealed partial class MusicTabViewModel : ViewModelBase
{
    public MusicTabViewModel()
    {
        SelectedService = MusicSettingsStore.SelectedService;
    }

    [ObservableProperty]
    public partial MusicService SelectedService { get; set; }

    partial void OnSelectedServiceChanged(MusicService value)
    {
        MusicSettingsStore.SelectedService = value;
        OnPropertyChanged(nameof(TabIconImage));
    }

    /// <summary>No text label needed — TabIconImage carries the active service's own logo, same idiom as the Whispers tab's 🤫 icon.</summary>
    public string Title => "";

    public IBrush HighlightBrush { get; } = new SolidColorBrush(Color.FromRgb(0x1E, 0xD7, 0x60));

    public Bitmap? TabIconImage => GameIconLoader.Get(SelectedService switch
    {
        MusicService.Spotify => "spotify",
        MusicService.Pandora => "pandora",
        _ => "youtube-music",
    });

    public double HeaderFontSize => 13;

    public IBrush HeaderForeground => HighlightBrush;

    /// <summary>No "something happened" concept for this tab — always false, matching a bot tab that's simply never marked unread.</summary>
    public bool HasUnread => false;
}
