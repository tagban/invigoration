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

    private async void OnAddBotClick(object? sender, RoutedEventArgs e)
    {
        var dialog = new ConfigWindow(new BotConfig());
        var result = await dialog.ShowDialog<BotConfig?>(this);
        if (result is not null)
        {
            ViewModel?.AddBot(result);
        }
    }

    private async void OnEditBotClick(object? sender, RoutedEventArgs e)
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

    private void OnOpenConfigFolderClick(object? sender, RoutedEventArgs e)
    {
        var dir = ConfigStore.DefaultConfigDirectory();
        Directory.CreateDirectory(dir);
        Process.Start(new ProcessStartInfo(dir) { UseShellExecute = true });
    }

    private void OnRemoveBotClick(object? sender, RoutedEventArgs e)
    {
        var vm = ViewModel;
        if (vm?.SelectedBot is { } selected)
        {
            vm.RemoveBot(selected);
        }
    }

    private async void OnAboutClick(object? sender, RoutedEventArgs e)
    {
        await new AboutWindow().ShowDialog(this);
    }

    private async void OnClanMembersClick(object? sender, RoutedEventArgs e)
    {
        await new ClanWindow().ShowDialog(this);
    }

    private async void OnClanRanksClick(object? sender, RoutedEventArgs e)
    {
        await new ClanRanksWindow().ShowDialog(this);
    }

    private async void OnManageIconsClick(object? sender, RoutedEventArgs e)
    {
        await new IconManagerWindow().ShowDialog(this);
    }

    private void OnExitClick(object? sender, RoutedEventArgs e)
    {
        if (Avalonia.Application.Current?.ApplicationLifetime is
            Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop)
        {
            desktop.Shutdown();
        }
    }
}
