using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Invigoration.App.Models;

namespace Invigoration.App.ViewModels;

/// <summary>
/// A pseudo bot-tab: sits in MainWindowViewModel.TopLevelTabs alongside the real BotTabViewModel
/// entries so the "Whispers" tab renders in the same TabControl/header style as every bot tab —
/// its ItemTemplate binds Title/HighlightBrush reflectively (no x:DataType pin), so any object
/// exposing those two properties slots in without needing type-based template selection.
/// Delegates its actual state to MainWindowViewModel rather than owning any, since
/// GlobalWhisperThreads already lives there (built by merging every bot's own WhisperThreads).
/// </summary>
public sealed class GlobalWhispersTabViewModel : ViewModelBase
{
    private readonly MainWindowViewModel _owner;

    public GlobalWhispersTabViewModel(MainWindowViewModel owner)
    {
        _owner = owner;
        // MainWindowViewModel owns the actual selection state (SelectedGlobalWhisperThread) —
        // re-raise here too whenever it changes, since a plain delegating property alone
        // wouldn't otherwise notify this object's own bindings (SelectedThread below).
        _owner.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName == nameof(MainWindowViewModel.SelectedGlobalWhisperThread))
            {
                OnPropertyChanged(nameof(SelectedThread));
            }
        };

        // HasUnread rolls up every thread's own flag — re-raise it whenever a thread's flag
        // changes, or a new thread (which starts unread) is added.
        Threads.CollectionChanged += OnThreadsCollectionChanged;
        foreach (var thread in Threads)
        {
            thread.PropertyChanged += OnThreadPropertyChanged;
        }
    }

    private void OnThreadsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.NewItems is not null)
        {
            foreach (WhisperThreadViewModel thread in e.NewItems)
            {
                thread.PropertyChanged += OnThreadPropertyChanged;
            }
        }

        if (e.OldItems is not null)
        {
            foreach (WhisperThreadViewModel thread in e.OldItems)
            {
                thread.PropertyChanged -= OnThreadPropertyChanged;
            }
        }

        OnPropertyChanged(nameof(HasUnread));
    }

    private void OnThreadPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(WhisperThreadViewModel.HasUnread))
        {
            OnPropertyChanged(nameof(HasUnread));
        }
    }

    /// <summary>No text label needed now that TabIconImage carries the "whisper" icon (defaults to 🤫, editable via Manage Icons) — an empty title just leaves the icon on its own in the tab header.</summary>
    public string Title => "";

    /// <summary>No bot-specific palette to draw from here — a fixed accent, distinct enough from the default Fluent accent to read as intentional.</summary>
    public IBrush HighlightBrush { get; } = new SolidColorBrush(Color.FromRgb(0x2C, 0xAC, 0xE8));

    /// <summary>The global Whispers tab's own icon — see IconManagerViewModel's "Custom Icons" section.</summary>
    public Bitmap? TabIconImage => GameIconLoader.Get("whisper");

    public double HeaderFontSize => 13;

    public IBrush HeaderForeground => HighlightBrush;

    /// <summary>True while any whisper thread across every bot has something unread — see the constructor's subscriptions.</summary>
    public bool HasUnread => Threads.Any(t => t.HasUnread);

    public ObservableCollection<WhisperThreadViewModel> Threads => _owner.GlobalWhisperThreads;

    public WhisperThreadViewModel? SelectedThread
    {
        get => _owner.SelectedGlobalWhisperThread;
        set => _owner.SelectedGlobalWhisperThread = value;
    }
}
