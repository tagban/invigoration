using System.Collections.ObjectModel;
using System.Diagnostics;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.Models;

namespace Invigoration.App.ViewModels;

/// <summary>
/// One icon in the picker. The picture is fetched the first time something asks for it, which
/// with a virtualized list means "when its row scrolls into view", so opening the picker doesn't
/// fire off 6,500 downloads.
/// </summary>
public sealed partial class HotlineIconCellViewModel(ushort id, HotlineIconInfo? info) : ObservableObject
{
    private bool _requested;
    private Bitmap? _image;

    public ushort Id { get; } = id;

    /// <summary>What the search index says this icon shows — null when the index isn't available.</summary>
    public HotlineIconInfo? Info { get; } = info;

    /// <summary>The icon's number, then its lettering, subjects and colors when the index has them.</summary>
    public string Tooltip => Info is null ? $"Icon {Id}" : $"Icon {Id}\n{Info.Describe()}";

    public Bitmap? Image
    {
        get
        {
            if (!_requested)
            {
                _requested = true;
                _ = LoadAsync();
            }

            return _image;
        }
    }

    private async Task LoadAsync()
    {
        _image = await HotlineIconLoader.GetAsync(Id).ConfigureAwait(true);
        OnPropertyChanged(nameof(Image));
    }
}

/// <summary>One line of the grid — a virtualizing panel only virtualizes rows, so the grid is built from rows of cells.</summary>
public sealed record HotlineIconRow(IReadOnlyList<HotlineIconCellViewModel> Cells);

/// <summary>
/// The icon chooser behind the session header's "Choose..." button: every icon hlwiki.com has,
/// searchable by number or by what's on it (hlwiki.com's ik0ns.csv — see HotlineIconIndex), with an optional "Download all" that fills the local cache so scrolling
/// is instant from then on.
/// </summary>
public sealed partial class HotlineIconPickerViewModel : ObservableObject
{
    private const int CellsPerRow = 4;

    /// <summary>A few at a time — it's one person's hobby site, not a CDN.</summary>
    private const int DownloadConcurrency = 4;

    private List<HotlineIconCellViewModel> _all = [];
    private IReadOnlyDictionary<ushort, HotlineIconInfo> _index = new Dictionary<ushort, HotlineIconInfo>();
    private CancellationTokenSource? _downloadCts;

    public ObservableCollection<HotlineIconRow> Rows { get; } = [];

    /// <summary>Set by the window — picking an icon closes it with that number.</summary>
    public Action<ushort>? Picked { get; set; }

    [ObservableProperty]
    public partial string FilterText { get; set; } = "";

    [ObservableProperty]
    public partial string StatusText { get; set; } = "Loading the icon list...";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(DownloadAllLabel))]
    public partial bool IsDownloading { get; set; }

    public string DownloadAllLabel => IsDownloading ? "Stop Downloading" : "Download All";

    /// <summary>Off each time the picker opens: icons the index tags "nsfw" stay out of the list unless asked for.</summary>
    [ObservableProperty]
    public partial bool ShowNsfw { get; set; }

    partial void OnFilterTextChanged(string value) => RebuildRows();

    partial void OnShowNsfwChanged(bool value) => RebuildRows();

    public async Task LoadAsync()
    {
        _index = await HotlineIconIndex.LoadAsync().ConfigureAwait(true);
        var ids = HotlineIconLoader.CatalogFrom(_index);
        _all = [.. ids.Select(id => new HotlineIconCellViewModel(id, _index.GetValueOrDefault(id)))];
        RebuildRows();
        StatusText = _all.Count == 0
            ? "Couldn't reach hlwiki.com for the icon list — type a number in the Icon box instead."
            : _index.Count == 0
                ? $"{_all.Count:N0} icons. Scroll, or type a number to filter."
                : $"{_all.Count:N0} icons. Scroll, or search by number, words on the icon, what's on it, or color.";
    }

    private void RebuildRows()
    {
        var filter = FilterText.Trim();
        var words = HotlineIconInfo.Split(filter).ToList();
        var allowed = ShowNsfw ? _all : _all.Where(c => c.Info is not { IsNsfw: true }).ToList();
        var matching = filter.Length == 0 || (words.Count == 0 && !filter.All(char.IsAsciiDigit))
            ? allowed
            : allowed.Where(c => HotlineIconIndex.Matches(c.Id, c.Info, words, filter)).ToList();
        Rows.Clear();
        for (var i = 0; i < matching.Count; i += CellsPerRow)
        {
            Rows.Add(new HotlineIconRow(matching.Skip(i).Take(CellsPerRow).ToList()));
        }

        if (filter.Length > 0 && _all.Count > 0)
        {
            StatusText = $"{matching.Count:N0} of {_all.Count:N0} icons match.";
        }
    }

    [RelayCommand]
    private void Pick(HotlineIconCellViewModel cell) => Picked?.Invoke(cell.Id);

    [RelayCommand]
    private static void OpenGallery() => Process.Start(new ProcessStartInfo(HotlineIconLoader.GalleryUrl) { UseShellExecute = true });

    [RelayCommand]
    private async Task ToggleDownloadAllAsync()
    {
        if (IsDownloading)
        {
            _downloadCts?.Cancel();
            return;
        }

        if (await TryDownloadZipAsync().ConfigureAwait(true))
        {
            return;
        }

        var missing = _all.Where(c => !HotlineIconLoader.IsCached(c.Id)).Select(c => c.Id).ToList();
        if (missing.Count == 0)
        {
            StatusText = $"All {_all.Count:N0} icons are already downloaded.";
            return;
        }

        IsDownloading = true;
        _downloadCts = new CancellationTokenSource();
        var ct = _downloadCts.Token;
        var done = 0;
        using var gate = new SemaphoreSlim(DownloadConcurrency);
        var tasks = missing.Select(async id =>
        {
            await gate.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                await HotlineIconLoader.GetAsync(id, ct).ConfigureAwait(false);
                var n = Interlocked.Increment(ref done);
                if (n % 25 == 0 || n == missing.Count)
                {
                    Dispatcher.UIThread.Post(() => StatusText = $"Downloading icons... {n:N0} of {missing.Count:N0}");
                }
            }
            finally
            {
                gate.Release();
            }
        });

        try
        {
            await Task.WhenAll(tasks).ConfigureAwait(true);
            StatusText = $"Downloaded {missing.Count:N0} icons — all {_all.Count:N0} are saved now.";
        }
        catch (OperationCanceledException)
        {
            StatusText = $"Stopped after {done:N0} of {missing.Count:N0}. What's downloaded stays saved.";
        }
        finally
        {
            IsDownloading = false;
        }
    }

    /// <summary>
    /// The whole set as one zip when the site has one — true when that worked (or was stopped),
    /// false to fall back to one icon at a time.
    /// </summary>
    private async Task<bool> TryDownloadZipAsync()
    {
        IsDownloading = true;
        _downloadCts = new CancellationTokenSource();
        StatusText = "Looking for the icon pack...";
        try
        {
            var saved = await HotlineIconLoader.DownloadZipAsync(
                new Progress<string>(text => StatusText = text), _downloadCts.Token).ConfigureAwait(true);
            if (saved is not { } count)
            {
                return false;
            }

            StatusText = $"Downloaded the icon pack — {count:N0} icons saved.";
            // Fresh cells, so any that had already come up empty read the newly saved files — from
            // the current index, so an icon taken off the site drops out of the list.
            var ids = HotlineIconLoader.CatalogFrom(_index);
            _all = [.. ids.Select(id => new HotlineIconCellViewModel(id, _index.GetValueOrDefault(id)))];
            RebuildRows();
            return true;
        }
        catch (OperationCanceledException)
        {
            StatusText = "Stopped. Download All starts the icon pack again.";
            return true;
        }
        catch (Exception ex) when (ex is HttpRequestException or InvalidDataException or IOException)
        {
            // A broken or unreachable zip shouldn't stop the icons from downloading at all.
            return false;
        }
        finally
        {
            IsDownloading = false;
        }
    }

    /// <summary>Closing the window stops a download in progress.</summary>
    public void Cancel() => _downloadCts?.Cancel();
}
