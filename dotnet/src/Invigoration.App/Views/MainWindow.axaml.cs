using System.Diagnostics;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Invigoration.App.ViewModels;
using Invigoration.Core;
using Invigoration.Core.Config;
using Invigoration.Core.Music;

namespace Invigoration.App.Views;

public partial class MainWindow : Window
{
    private const string BaseTitle = $"Invigoration v{AppVersion.Current}";

    private static readonly TimeSpan TitleUpdateInterval = TimeSpan.FromSeconds(15);

    public MainWindow()
    {
        InitializeComponent();
        Title = BaseTitle;
        Closing += (_, _) => ViewModel?.SaveAll();
        StartTitleUpdateTimer();
        // TopLevelTabs puts the Whispers pseudo-tab first, then (if enabled) Music — default to
        // the first real bot/group instead, matching this window's behavior before either
        // existed. A hardcoded SelectedIndex=1 (this used to be that) would land on Music instead
        // of a bot whenever Music is enabled, since it now also occupies an early slot. Loaded
        // (not the constructor) since DataContext isn't set yet at construction time.
        Loaded += (_, _) =>
        {
            if (this.FindControl<TabStrip>("TopLevelTabControl") is { } tabControl && ViewModel is { } vm)
            {
                var firstBotTab = vm.TopLevelTabs.FirstOrDefault(t => t is BotTabViewModel or BotGroupTabViewModel);
                if (firstBotTab is not null)
                {
                    tabControl.SelectedItem = firstBotTab;
                }

                vm.PropertyChanged += OnViewModelPropertyChanged;
            }
        };
    }

    /// <summary>
    /// Shows what's currently playing in the title bar (e.g. "Invigoration v2.0.3b - Spotify:
    /// Pink Pony Club") whenever the music player is open and something's playing, falling back
    /// to the plain BaseTitle otherwise. Polling (not event-driven) since GetNowPlayingAsync is
    /// pull-based — nothing raises an event when the track changes; DispatcherTimer (not a
    /// background Task/PeriodicTimer) since this only needs to run while the window exists and
    /// writes directly to a UI property.
    /// </summary>
    private void StartTitleUpdateTimer()
    {
        var timer = new DispatcherTimer { Interval = TitleUpdateInterval };
        timer.Tick += async (_, _) => await UpdateTitleAsync();
        timer.Start();
    }

    private async Task UpdateTitleAsync()
    {
        if (MusicPlayerRegistry.Controller is not { } controller)
        {
            Title = BaseTitle;
        }
        else
        {
            var nowPlaying = await controller.GetNowPlayingAsync();
            Title = nowPlaying is null ? BaseTitle : $"{BaseTitle} - {nowPlaying.Service}: {nowPlaying.Title}";
        }

        // Same tick also refreshes the optional bottom playback bar — no separate poll loop
        // needed for it (see MusicBarViewModel's remarks).
        if (ViewModel is { IsMusicBarEnabled: true } vm)
        {
            await vm.MusicBar.RefreshAsync();
        }
    }

    /// <summary>
    /// Toggling the Customize menu's "Music Player" checkbox rebuilds TopLevelTabs (RefreshTopLevelTabs), which — same reset-to-first-item issue AddBot/EditBot already had to work around via SelectTopLevelBot — would otherwise leave the tab strip showing Whispers instead of the Music tab the user just turned on.
    /// Also mirrors MainWindowViewModel.FocusWhisperThread (the right-click "Whisper" action) onto the actual TabStrip control — SelectedGlobalWhisperThread itself can't drive tab selection directly since TopLevelTabs' selection isn't bound TwoWay (see SelectedTopLevelItem's remarks).
    /// </summary>
    private void OnViewModelPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (ViewModel is not { } vm || this.FindControl<TabStrip>("TopLevelTabControl") is not { } tabControl)
        {
            return;
        }

        if (e.PropertyName == nameof(MainWindowViewModel.IsMusicEnabled) && vm.IsMusicEnabled)
        {
            tabControl.SelectedItem = vm.MusicTab;
        }
        else if (e.PropertyName == nameof(MainWindowViewModel.SelectedGlobalWhisperThread) && vm.SelectedGlobalWhisperThread is not null)
        {
            tabControl.SelectedItem = vm.WhispersTab;
        }
    }

    private MainWindowViewModel? ViewModel => DataContext as MainWindowViewModel;

    /// <summary>Keeps MainWindowViewModel.SelectedBot in sync with whichever tab is actually showing — TabControl.SelectedItem can't bind directly to it two-way any more since TopLevelTabs is a mixed BotTabViewModel/GlobalWhispersTabViewModel collection; selecting the Whispers tab leaves SelectedBot as whatever bot was last actually selected, which is a reasonable "last bot you were looking at" fallback for the Edit/Remove Selected Bot menu actions. Also feeds SetActiveTopLevelItem, which drives every bot's IsActive/HasUnread state (see RecomputeActiveBot).</summary>
    private void OnTopLevelTabSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (ViewModel is not { } vm)
        {
            return;
        }

        var selected = e.AddedItems.Count > 0 ? e.AddedItems[0] : null;
        vm.SetActiveTopLevelItem(selected);
        if (selected is BotTabViewModel bot)
        {
            vm.SelectedBot = bot;
        }
    }

    // Each action below has two Click handlers with identical bodies: one typed for the
    // in-window Menu (RoutedEventArgs) and one for the macOS NativeMenu (plain EventArgs).
    // Avalonia's XAML compiler requires an exact delegate match, so the two can't share a
    // single method despite RoutedEventArgs being an EventArgs.

    private async void OnAddBotClick(object? sender, RoutedEventArgs e) => await AddBot();
    private async void OnAddBotNativeClick(object? sender, EventArgs e) => await AddBot();

    private async Task AddBot()
    {
        var dialog = new ConfigWindow(new BotConfig());
        var result = await dialog.ShowDialog<BotConfig?>(this);
        if (result is not null && ViewModel is { } vm)
        {
            vm.AddBot(result);
            // AddBot's own RefreshTopLevelTabs() call clears and repopulates TopLevelTabs, which
            // resets the TabControl's own selection to index 0 (the Whispers pseudo-tab, always
            // first) — SelectedBot is a separate ViewModel property, not something the TabControl
            // actually reads. Restore it to the bot AddBot just set SelectedBot to.
            if (vm.SelectedBot is { } added)
            {
                SelectTopLevelBot(added);
            }
        }
    }

    private async void OnEditBotClick(object? sender, RoutedEventArgs e) => await EditSelectedBot();
    private async void OnEditBotNativeClick(object? sender, EventArgs e) => await EditSelectedBot();

    private async Task EditSelectedBot()
    {
        var vm = ViewModel;
        if (vm?.SelectedBot is not { } selected)
        {
            return;
        }

        var dialog = new ConfigWindow(selected.Config);
        var result = await dialog.ShowDialog<BotConfig?>(this);
        if (result is not null)
        {
            selected.ApplyConfig(result);
            vm.RefreshTopLevelTabs();
            vm.SaveAll();
            // Same TabControl-selection-reset issue as AddBot — RefreshTopLevelTabs() rebuilding
            // TopLevelTabs from scratch otherwise leaves the tab strip showing Whispers instead of
            // the bot that was actually just edited.
            SelectTopLevelBot(selected);
        }
    }

    /// <summary>Selects a bot's own top-level tab directly if it's ungrouped, or its containing group's tab (and that bot within the group's own nested TabControl) if it's now grouped — used after RefreshTopLevelTabs() rebuilds TopLevelTabs and resets the TabControl's own selection.</summary>
    private void SelectTopLevelBot(BotTabViewModel bot)
    {
        if (this.FindControl<TabStrip>("TopLevelTabControl") is not { } tabControl || ViewModel is not { } vm)
        {
            return;
        }

        if (vm.TopLevelTabs.Contains(bot))
        {
            tabControl.SelectedItem = bot;
            return;
        }

        var group = vm.TopLevelTabs.OfType<BotGroupTabViewModel>().FirstOrDefault(g => g.Bots.Contains(bot));
        if (group is null)
        {
            return;
        }

        group.SelectedBot = bot;
        tabControl.SelectedItem = group;
    }

    private void OnOpenConfigFolderClick(object? sender, RoutedEventArgs e) => OpenConfigFolder();
    private void OnOpenConfigFolderNativeClick(object? sender, EventArgs e) => OpenConfigFolder();

    private static void OpenConfigFolder()
    {
        var dir = ConfigStore.DefaultConfigDirectory();
        Directory.CreateDirectory(dir);
        Process.Start(new ProcessStartInfo(dir) { UseShellExecute = true });
    }

    private void OnRemoveBotClick(object? sender, RoutedEventArgs e) => RemoveSelectedBot();
    private void OnRemoveBotNativeClick(object? sender, EventArgs e) => RemoveSelectedBot();

    private void RemoveSelectedBot()
    {
        var vm = ViewModel;
        if (vm?.SelectedBot is { } selected)
        {
            vm.RemoveBot(selected);
        }
    }

    private async void OnAboutClick(object? sender, RoutedEventArgs e) => await ShowAbout();
    private async void OnAboutNativeClick(object? sender, EventArgs e) => await ShowAbout();

    private async Task ShowAbout() => await new AboutWindow().ShowDialog(this);

    private async void OnClanMembersClick(object? sender, RoutedEventArgs e) => await ShowClanMembers();
    private async void OnClanMembersNativeClick(object? sender, EventArgs e) => await ShowClanMembers();

    private async Task ShowClanMembers() => await new ClanWindow().ShowDialog(this);

    private async void OnClanRanksClick(object? sender, RoutedEventArgs e) => await ShowClanRanks();
    private async void OnClanRanksNativeClick(object? sender, EventArgs e) => await ShowClanRanks();

    private async Task ShowClanRanks() => await new ClanRanksWindow().ShowDialog(this);

    private async void OnManageIconsClick(object? sender, RoutedEventArgs e) => await ShowIconManager();
    private async void OnManageIconsNativeClick(object? sender, EventArgs e) => await ShowIconManager();

    private async Task ShowIconManager() => await new IconManagerWindow().ShowDialog(this);

    private async void OnManageColorsClick(object? sender, RoutedEventArgs e) => await ShowColorManager();
    private async void OnManageColorsNativeClick(object? sender, EventArgs e) => await ShowColorManager();

    private async Task ShowColorManager() => await new ColorManagerWindow().ShowDialog(this);

    private async void OnManageBattlenetProfilesClick(object? sender, RoutedEventArgs e) => await ShowBattlenetProfiles();
    private async void OnManageBattlenetProfilesNativeClick(object? sender, EventArgs e) => await ShowBattlenetProfiles();

    private async Task ShowBattlenetProfiles() => await new BattlenetCredentialProfilesWindow().ShowDialog(this);

    private void OnExitClick(object? sender, RoutedEventArgs e) => ExitApplication();
    private void OnExitNativeClick(object? sender, EventArgs e) => ExitApplication();

    private static void ExitApplication()
    {
        if (Avalonia.Application.Current?.ApplicationLifetime is
            Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop)
        {
            desktop.Shutdown();
        }
    }
}
