using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Diagnostics;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.App.Models;
using Invigoration.Core;
using Invigoration.Core.Config;
using Invigoration.Core.Hotline;
using Invigoration.Core.Music;
using Invigoration.Core.Trivia;

namespace Invigoration.App.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    private readonly ConfigStore _store = new();

    public ObservableCollection<BotTabViewModel> Bots { get; } = [];

    [ObservableProperty]
    public partial BotTabViewModel? SelectedBot { get; set; }

    /// <summary>
    /// What MainWindow's top-level TabControl actually binds ItemsSource to — the Whispers
    /// pseudo-tab first, then one entry per real bot. Deliberately typed as plain `object`, not
    /// `ViewModelBase`: Avalonia's XAML compiler infers a compiled-binding type from an
    /// ObservableCollection's generic parameter, and a strongly-typed ViewModelBase would make
    /// the header ItemTemplate's Title/HighlightBrush bindings fail to compile (ViewModelBase
    /// itself has neither) — `object` forces reflection-based binding instead, which is exactly
    /// what lets GlobalWhispersTabViewModel duck-type those same two property names and render
    /// in the same header template as a real BotTabViewModel. Kept in sync with Bots wherever it
    /// changes (constructor, AddBot, RemoveBot) — no separate CollectionChanged bridging needed
    /// since Bots itself is never reordered after add.
    /// </summary>
    public ObservableCollection<object> TopLevelTabs { get; } = [];

    /// <summary>
    /// Every bot's WhisperThreads merged into one most-recently-active-first list, for the
    /// top-level Whispers tab that spans all connected bots — see WireGlobalWhispers. Each
    /// thread is the exact same instance its own bot's per-bot Whispers tab shows (via
    /// WhisperThreadViewModel.Owner), so replying/marking read from either place stays in sync.
    /// </summary>
    public ObservableCollection<WhisperThreadViewModel> GlobalWhisperThreads { get; } = [];

    [ObservableProperty]
    public partial WhisperThreadViewModel? SelectedGlobalWhisperThread { get; set; }

    partial void OnSelectedGlobalWhisperThreadChanged(WhisperThreadViewModel? value) => value?.MarkRead();

    /// <summary>
    /// Icon lookup has no per-bot concept — IconOverrideStore is one shared
    /// folder, and IconSetStore.ApplySet swaps its contents wholesale. Since
    /// each bot can now name its own preferred set (BotConfig.IconSetName),
    /// the closest approximation to "per-bot" without threading a set name
    /// through every icon lookup call site is re-applying the newly-selected
    /// bot's set whenever the tab selection changes — icons are then correct
    /// for whichever bot you're actually looking at, even though two tabs
    /// can't render different sets *simultaneously*.
    /// </summary>
    partial void OnSelectedBotChanged(BotTabViewModel? value)
    {
        if (!string.IsNullOrEmpty(value?.Config.IconSetName))
        {
            IconSetStore.ApplySet(value.Config.IconSetName);
        }
    }

    /// <summary>
    /// True once at least one configured bot has ClanFeatureEnabled on — the
    /// top-level "Clan" menu binds its IsVisible to this so it disappears
    /// entirely rather than offering clan management nobody's turned on.
    /// Recomputed wherever bot config can change (load, add, remove, and
    /// SaveAll after an edit) since BotConfig itself isn't observable.
    /// </summary>
    [ObservableProperty]
    public partial bool AnyBotHasClanEnabled { get; set; }

    private readonly GlobalWhispersTabViewModel _whispersTab;

    /// <summary>Exposed so a context-menu "Whisper" click deep inside a bot's own BotTabView (see BotTabView.axaml.cs's FocusWhisperThread) can switch the top-level tab strip to Whispers without needing its own reference to this pseudo-tab.</summary>
    public GlobalWhispersTabViewModel WhispersTab => _whispersTab;

    /// <summary>Exposed directly (not just via TopLevelTabs) so MainWindow.axaml can bind the permanent MusicPlayerPanel overlay's DataContext to the exact same instance the tab header represents — see MusicPlayerPanel's remarks for why it lives outside the TabControl's own content area.</summary>
    public MusicTabViewModel MusicTab { get; } = new();

    /// <summary>Whether the Music tab (and its always-alive player overlay) shows at all — the Customize menu's "Music Player" toggle, for anyone who doesn't want it. Persisted via MusicSettingsStore, not per-bot.</summary>
    [ObservableProperty]
    public partial bool IsMusicEnabled { get; set; }

    partial void OnIsMusicEnabledChanged(bool value)
    {
        MusicSettingsStore.IsEnabled = value;
        RefreshTopLevelTabs();
    }

    /// <summary>True exactly when the Music tab's header is the one currently showing — drives MusicPlayerPanel's IsVisible in MainWindow.axaml, since the panel itself is never added/removed from the tree (see its remarks).</summary>
    [ObservableProperty]
    public partial bool IsMusicTabSelected { get; set; }

    /// <summary>The optional bottom playback-control bar's own state/commands — see MusicBarViewModel's remarks. Always exists (like MusicTab); IsMusicBarEnabled just controls whether MainWindow.axaml actually shows it.</summary>
    public MusicBarViewModel MusicBar { get; } = new();

    /// <summary>Off by default — a thin persistent playback bar docked at the bottom of the whole window, independent of whether the Music tab itself is enabled/selected. Toggled via the Customize menu.</summary>
    [ObservableProperty]
    public partial bool IsMusicBarEnabled { get; set; }

    partial void OnIsMusicBarEnabledChanged(bool value) => MusicSettingsStore.ShowBottomBar = value;

    private readonly HotlineTrackerConfigStore _hotlineStore = new();

    /// <summary>"Add Bot"-style top-level Hotline entities, per the user's own framing ("each hotline connection is a 'tracker'") — one HotlineTabViewModel per HotlineTrackerConfig, persisted the same way Bots/BotConfig are.</summary>
    public ObservableCollection<HotlineTabViewModel> HotlineTrackers { get; } = [];

    public MainWindowViewModel()
    {
        _whispersTab = new GlobalWhispersTabViewModel(this);
        IsMusicEnabled = MusicSettingsStore.IsEnabled;
        IsMusicBarEnabled = MusicSettingsStore.ShowBottomBar;

        foreach (var config in _store.Load())
        {
            var tab = CreateBotTab(config);
            Bots.Add(tab);
            WireGlobalWhispers(tab);
        }

        foreach (var config in _hotlineStore.Load())
        {
            HotlineTrackers.Add(CreateHotlineTab(config));
        }

        RefreshTopLevelTabs();
        SelectedBot = Bots.Count > 0 ? Bots[0] : null;
        RefreshAnyBotHasClanEnabled();
        _ = AutoConnectStartupBotsAsync();
        foreach (var tracker in HotlineTrackers)
        {
            tracker.Tracker.AutoConnectStartupProfiles();
        }

        // Fire-and-forget and best-effort: seeds the Trivia folder with the base packs from
        // GitHub the first time (see TriviaPackDownloader), never blocks startup, and is a
        // no-op on every later launch once those files already exist locally.
        _ = TriviaPackDownloader.EnsureDownloadedAsync(err => Debug.WriteLine(err));
    }

    private HotlineTabViewModel CreateHotlineTab(HotlineTrackerConfig config)
    {
        var tab = new HotlineTabViewModel(config);
        tab.ConfigChanged += SaveHotlineTrackers;
        tab.RemoveRequested += () => RemoveHotlineTracker(tab);
        return tab;
    }

    /// <summary>The Hotline equivalent of AddBot — creates a new tracker with sensible defaults (renamed/re-hosted in place afterward, same in-place-editing idiom as a saved server profile, rather than a separate dialog). Returns the new tab so MainWindow.axaml.cs can select it on the actual TabStrip control, same as AddBot/SelectTopLevelBot.</summary>
    public HotlineTabViewModel AddHotlineTracker()
    {
        var name = HotlineTrackers.Count == 0 ? "Hotline" : $"Hotline {HotlineTrackers.Count + 1}";
        var config = new HotlineTrackerConfig { DisplayName = name };
        var tab = CreateHotlineTab(config);
        HotlineTrackers.Add(tab);
        SaveHotlineTrackers();
        RefreshTopLevelTabs();
        return tab;
    }

    public void RemoveHotlineTracker(HotlineTabViewModel tab)
    {
        HotlineTrackers.Remove(tab);
        SaveHotlineTrackers();
        RefreshTopLevelTabs();
    }

    private void SaveHotlineTrackers() => _hotlineStore.Save(HotlineTrackers.Select(t => t.Config).ToList());

    private void RefreshAnyBotHasClanEnabled() => AnyBotHasClanEnabled = Bots.Any(b => b.Config.ClanFeatureEnabled);

    /// <summary>
    /// Rebuilds TopLevelTabs from Bots' current BotConfig.TabGroup values — the Whispers
    /// pseudo-tab first, then every ungrouped bot as its own tab, then one BotGroupTabViewModel
    /// per distinct non-empty TabGroup. Deliberately NOT tied to every SaveAll() (that fires
    /// routinely from unrelated background activity, e.g. an SC2 channel join persisting
    /// Sc2LastChannelNames — rebuilding the tab strip that often would reset selection/scroll
    /// state for no reason); called explicitly instead from the few places TabGroup can actually
    /// change: startup, AddBot, RemoveBot, and after the Config window closes with changes.
    /// </summary>
    public void RefreshTopLevelTabs()
    {
        TopLevelTabs.Clear();
        TopLevelTabs.Add(_whispersTab);
        if (IsMusicEnabled)
        {
            TopLevelTabs.Add(MusicTab);
        }

        foreach (var tracker in HotlineTrackers)
        {
            TopLevelTabs.Add(tracker);
        }

        foreach (var bot in Bots.Where(b => string.IsNullOrEmpty(b.Config.TabGroup)))
        {
            TopLevelTabs.Add(bot);
        }

        foreach (var group in Bots.Where(b => !string.IsNullOrEmpty(b.Config.TabGroup))
                     .GroupBy(b => b.Config.TabGroup)
                     .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
        {
            var groupTab = new BotGroupTabViewModel(group.Key, group);
            // A click on one of this group's own nested sub-tabs changes SelectedBot here
            // without the outer TabControl's SelectionChanged ever firing — recompute so that
            // bot's IsActive/HasUnread reflect it too, but only when this group is actually the
            // visible top-level tab right now (RecomputeActiveBot itself checks that).
            groupTab.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(BotGroupTabViewModel.SelectedBot))
                {
                    RecomputeActiveBot();
                }
            };
            TopLevelTabs.Add(groupTab);
        }

        RecomputeActiveBot();
    }

    /// <summary>
    /// Whichever top-level item is currently selected — bound by MainWindow.axaml's content-area
    /// ContentControl (TabStrip no longer manages content itself; see that XAML's remarks) as
    /// well as read here for RecomputeActiveBot. Set from MainWindow.axaml.cs's
    /// OnTopLevelTabSelectionChanged via SetActiveTopLevelItem below, not bound directly
    /// TwoWay to TabStrip.SelectedItem, to keep the existing event-driven wiring intact.
    /// </summary>
    [ObservableProperty]
    public partial object? SelectedTopLevelItem { get; set; }

    /// <summary>Called by MainWindow.axaml.cs whenever the top-level TabStrip's selection changes — records which item is showing and recomputes which bot (if any) is actually visible.</summary>
    public void SetActiveTopLevelItem(object? item)
    {
        SelectedTopLevelItem = item;
        IsMusicTabSelected = ReferenceEquals(item, MusicTab);
        RecomputeActiveBot();
    }

    /// <summary>
    /// The right-click "Whisper" action on a bot's classic Users list (BotTabView.axaml.cs) goes
    /// through here rather than just setting whisper-focus silently: finds/creates the thread for
    /// that peer, selects it (so it's ready to reply to the instant the Whispers tab shows), and
    /// switches the top-level tab strip there — MainWindow.axaml.cs mirrors SelectedGlobalWhisperThread
    /// changes onto the actual TabStrip control (it can't be bound TwoWay directly, see SelectedTopLevelItem's remarks).
    /// </summary>
    public void FocusWhisperThread(BotTabViewModel bot, string peer) =>
        SelectedGlobalWhisperThread = bot.GetOrCreateWhisperThread(peer);

    /// <summary>
    /// "Active" (IsActive/HasUnread-clearing) means genuinely visible right now — for a plain
    /// bot tab that's just itself; for a group tab, it's whichever bot the group's own nested
    /// TabControl currently has selected, not the group as a whole. Re-run whenever either
    /// selection level changes (see SetActiveTopLevelItem and the groupTab.PropertyChanged
    /// subscription above) so both stay in sync with what's actually on screen.
    ///
    /// Also keeps SelectedBot itself pointed at whichever bot is genuinely visible — this used
    /// to be MainWindow.axaml.cs's job alone (OnTopLevelTabSelectionChanged), but that only ever
    /// fires for a plain top-level bot tab, never for switching sub-tabs *inside* a group's own
    /// nested TabControl. A grouped bot's own click only ever reached this method (via the
    /// groupTab.PropertyChanged subscription below), so SelectedBot — what "Edit/Remove Selected
    /// Bot" actually reads — stayed frozen on whichever bot last selected it directly (e.g. the
    /// last-added bot), regardless of which grouped sub-tab was actually showing. Skipped when
    /// active is null (e.g. the Whispers tab is showing) so the existing "last real bot you were
    /// looking at" fallback for those menu actions is preserved.
    /// </summary>
    private void RecomputeActiveBot()
    {
        var active = SelectedTopLevelItem switch
        {
            BotTabViewModel bot => bot,
            BotGroupTabViewModel group => group.SelectedBot,
            _ => null,
        };

        foreach (var bot in Bots)
        {
            bot.IsActive = bot == active;
        }

        if (active is not null)
        {
            SelectedBot = active;
        }
    }

    /// <summary>
    /// Connects every bot with AutoConnectOnStartup, staggered a couple
    /// seconds apart rather than all at once — several bots opening TCP
    /// connections to the same server in the same instant is exactly the
    /// kind of burst a per-IP connection or flood limit can catch.
    /// </summary>
    private async Task AutoConnectStartupBotsAsync()
    {
        foreach (var bot in Bots.Where(b => b.Config.AutoConnectOnStartup).ToList())
        {
            await bot.ConnectCommand.ExecuteAsync(null);
            await Task.Delay(2000);
        }
    }

    /// <summary>Wires BotEngine.ConfigPersistNeeded (see its own remarks) to SaveAll, so a config mutation the engine makes on its own — currently just auto-assigning a Battle.net credential profile on first SC2 connect — reaches bots.json rather than only living in memory.</summary>
    private BotTabViewModel CreateBotTab(BotConfig config)
    {
        var engine = new BotEngine(config);
        engine.ConfigPersistNeeded += SaveAll;
        return new BotTabViewModel(engine);
    }

    public void AddBot(BotConfig config)
    {
        var tab = CreateBotTab(config);
        Bots.Add(tab);
        WireGlobalWhispers(tab);
        RefreshTopLevelTabs();
        SelectedBot = tab;
        SaveAll();
    }

    public async void RemoveBot(BotTabViewModel tab)
    {
        Bots.Remove(tab);
        // A straight TopLevelTabs.Remove(tab) wouldn't reach a grouped bot at all — it isn't a
        // direct member of TopLevelTabs, it's nested inside its BotGroupTabViewModel.Bots — so
        // this rebuilds instead, correctly dropping it whether it was grouped or not.
        RefreshTopLevelTabs();
        if (SelectedBot == tab)
        {
            SelectedBot = Bots.Count > 0 ? Bots[0] : null;
        }

        UnwireGlobalWhispers(tab);
        tab.Engine.ConfigPersistNeeded -= SaveAll;
        await tab.DisposeAsync();
        SaveAll();
    }

    /// <summary>
    /// Mirrors one bot's WhisperThreads into GlobalWhisperThreads: a new thread is inserted at
    /// the top (matches per-bot ordering, most-recently-active-first), and a Move — fired by
    /// UpsertWhisper bumping an existing thread back to the top on new activity — mirrors the
    /// same reordering here, so the global tab's ordering stays "most recently active across
    /// every bot" too, not just within each bot's own list.
    /// </summary>
    private void WireGlobalWhispers(BotTabViewModel bot)
    {
        foreach (var thread in bot.WhisperThreads)
        {
            GlobalWhisperThreads.Insert(0, thread);
        }

        bot.WhisperThreads.CollectionChanged += OnBotWhisperThreadsChanged;
    }

    private void UnwireGlobalWhispers(BotTabViewModel bot)
    {
        bot.WhisperThreads.CollectionChanged -= OnBotWhisperThreadsChanged;
        foreach (var thread in bot.WhisperThreads)
        {
            GlobalWhisperThreads.Remove(thread);
        }
    }

    private void OnBotWhisperThreadsChanged(object? sender, NotifyCollectionChangedEventArgs e) => Dispatcher.UIThread.Post(() =>
    {
        switch (e.Action)
        {
            case NotifyCollectionChangedAction.Add when e.NewItems is not null:
                foreach (WhisperThreadViewModel thread in e.NewItems)
                {
                    GlobalWhisperThreads.Insert(0, thread);
                }

                break;

            case NotifyCollectionChangedAction.Move when e.NewItems is not null:
                foreach (WhisperThreadViewModel thread in e.NewItems)
                {
                    var index = GlobalWhisperThreads.IndexOf(thread);
                    if (index > 0)
                    {
                        GlobalWhisperThreads.Move(index, 0);
                    }
                }

                break;

            case NotifyCollectionChangedAction.Remove when e.OldItems is not null:
                foreach (WhisperThreadViewModel thread in e.OldItems)
                {
                    GlobalWhisperThreads.Remove(thread);
                }

                break;
        }
    });

    public void SaveAll()
    {
        _store.Save(Bots.Select(b => b.Config).ToList());
        RefreshAnyBotHasClanEnabled();
    }
}
