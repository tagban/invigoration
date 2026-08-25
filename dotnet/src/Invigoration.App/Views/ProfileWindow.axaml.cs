using Avalonia.Controls;
using Avalonia.Interactivity;
using Invigoration.App.ViewModels;
using Invigoration.Core;

namespace Invigoration.App.Views;

public partial class ProfileWindow : Window
{
    public ProfileWindow(BotEngine engine, string account)
    {
        InitializeComponent();
        DataContext = new ProfileViewModel(engine, account);
    }

    private void OnCloseClick(object? sender, RoutedEventArgs e) => Close();
}
