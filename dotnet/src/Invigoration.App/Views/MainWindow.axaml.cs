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
    }

    private MainWindowViewModel? ViewModel => DataContext as MainWindowViewModel;

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
