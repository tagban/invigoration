using Avalonia.Controls;
using Avalonia.Interactivity;
using Invigoration.App.ViewModels;

namespace Invigoration.App.Views;

public partial class BattlenetCredentialProfilesWindow : Window
{
    private readonly BattlenetCredentialProfilesViewModel _viewModel = new();

    public BattlenetCredentialProfilesWindow()
    {
        InitializeComponent();
        DataContext = _viewModel;
    }

    private async void OnSignInClick(object? sender, RoutedEventArgs e)
    {
        if ((sender as Button)?.DataContext is not BattlenetCredentialProfileViewModel profile)
        {
            return;
        }

        try
        {
            await profile.SignInAsync(this);
        }
        catch (Exception ex)
        {
            _viewModel.StatusMessage = ex.Message;
        }
    }

    private void OnSaveClick(object? sender, RoutedEventArgs e) => _viewModel.SaveCommand.Execute(null);

    private void OnCloseClick(object? sender, RoutedEventArgs e) => Close();
}
