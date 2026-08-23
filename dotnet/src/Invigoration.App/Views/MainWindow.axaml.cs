using System.Diagnostics;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Invigoration.App.ViewModels;
using Invigoration.Core;
using Invigoration.Core.Config;

namespace Invigoration.App.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        Title = $"Invigoration v{AppVersion.Current}";
        Closing += (_, _) => ViewModel?.SaveAll();
        // TopLevelTabs puts the Whispers pseudo-tab first (index 0) — default to the first real
        // bot instead, matching this window's behavior before that tab existed. Loaded (not the
        // constructor) since DataContext isn't set yet at construction time.
        Loaded += (_, _) =>
        {
            if (this.FindControl<TabControl>("TopLevelTabControl") is { } tabControl && ViewModel is { Bots.Count: > 0 })
            {
                tabControl.SelectedIndex = 1;
            }
        };
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
        if (result is not null)
        {
            ViewModel?.AddBot(result);
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
        }
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
