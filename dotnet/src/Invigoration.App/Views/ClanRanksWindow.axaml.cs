using Avalonia.Controls;
using Avalonia.Interactivity;
using Invigoration.App.ViewModels;

namespace Invigoration.App.Views;

public partial class ClanRanksWindow : Window
{
    private readonly ClanRanksViewModel _viewModel = new();

    public ClanRanksWindow()
    {
        InitializeComponent();
        DataContext = _viewModel;
    }

    private void OnSaveClick(object? sender, RoutedEventArgs e) => _viewModel.SaveCommand.Execute(null);

    private void OnCloseClick(object? sender, RoutedEventArgs e) => Close();
}
