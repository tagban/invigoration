using System.Collections.ObjectModel;
using Avalonia.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.Core.Hotline;

namespace Invigoration.App.ViewModels;

/// <summary>
/// One "Add Bot"-style top-level Hotline entity (see MainWindowViewModel.HotlineTrackers) —
/// wraps a HotlineTrackerConfig the same way BotTabViewModel wraps a BotConfig, duck-typed into
/// the same TabStrip header template via Title/HighlightBrush/HeaderFontSize/HeaderForeground/
/// HasUnread/TabIconImage.
///
/// Its own content is a nested TabControl (see HotlineTabView.axaml, the same pattern
/// BotGroupTabViewModel uses for its per-bot sub-tabs) whose first, permanent entry is the
/// tracker/saved-profiles browser (HotlineTrackerViewModel — "the primary window starts with just
/// the tracker", per the original request) and whose later entries are one HotlineSessionViewModel
/// per server actually connected to — "a group for tabs of servers".
/// </summary>
public sealed partial class HotlineTabViewModel : ViewModelBase
{
    /// <summary>Shared accent for the outer tab's underline and every session sub-tab's underline — Hotline has no per-server palette of its own to draw from, same reasoning as BotGroupTabViewModel borrowing a color rather than inventing one per group.</summary>
    public static readonly IBrush AccentBrush = new SolidColorBrush(Color.FromRgb(0x3B, 0x9A, 0xE8));

    public HotlineTrackerConfig Config { get; }

    public string Title => Config.DisplayName;
    public IBrush HighlightBrush => AccentBrush;
    public double HeaderFontSize => 13;
    public IBrush HeaderForeground => Brushes.White;

    /// <summary>Extracted directly from the user's own bigredh.com title bar art (its red "H" mark), background-removed the same way the other bundled icon sets were.</summary>
    public Avalonia.Media.Imaging.Bitmap? TabIconImage => Invigoration.App.Models.GameIconLoader.Get("hotline");

    public bool HasUnread => Items.OfType<HotlineSessionViewModel>().Any(s => s.HasUnread);

    public ObservableCollection<object> Items { get; } = [];

    public HotlineTrackerViewModel Tracker { get; }

    [ObservableProperty]
    public partial object? SelectedItem { get; set; }

    /// <summary>Set by MainWindowViewModel — a "Remove Tracker" click (see HotlineTrackerView) goes through this rather than the view model reaching back into MainWindowViewModel directly, same event-callback shape as BotEngine.ConfigPersistNeeded.</summary>
    public event Action? RemoveRequested;

    /// <summary>Fired whenever Config's own fields change (name/host/agreement setting edited in HotlineTrackerView) — MainWindowViewModel subscribes to persist the whole tracker list, same event-callback shape as BotEngine.ConfigPersistNeeded.</summary>
    public event Action? ConfigChanged;

    public HotlineTabViewModel(HotlineTrackerConfig config)
    {
        Config = config;
        Tracker = new HotlineTrackerViewModel(this);
        Items.Add(Tracker);
        SelectedItem = Tracker;
    }

    partial void OnSelectedItemChanged(object? value)
    {
        // Arriving on a session's own sub-tab is what actually clears its unread flag — matches
        // BotTabViewModel's own "being looked at clears unread" idiom.
        if (value is HotlineSessionViewModel session)
        {
            session.HasUnread = false;
        }
        else if (value is HotlineTrackerViewModel tracker)
        {
            // Refreshes the tracker's server list every time you land on it — including its very
            // first appearance, since this constructor's own SelectedItem = Tracker below fires
            // this same partial method — per explicit request, instead of requiring a manual
            // Refresh click.
            tracker.OnTabActivated();
        }

        OnPropertyChanged(nameof(HasUnread));
    }

    public void Connect(HotlineConnectOptions options)
    {
        var session = new HotlineSessionViewModel(this, options);
        session.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName == nameof(HotlineSessionViewModel.HasUnread))
            {
                OnPropertyChanged(nameof(HasUnread));
            }
        };
        Items.Add(session);
        SelectedItem = session;
    }

    public void CloseSession(HotlineSessionViewModel session)
    {
        Items.Remove(session);
        if (SelectedItem == session)
        {
            SelectedItem = Tracker;
        }
    }

    public void RequestRemove() => RemoveRequested?.Invoke();

    public void NotifyConfigChanged()
    {
        OnPropertyChanged(nameof(Title));
        ConfigChanged?.Invoke();
    }
}
