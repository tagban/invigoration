using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Invigoration.App.ViewModels;

namespace Invigoration.App.Views;

public partial class ThemeManagerWindow : Window
{
    private readonly ThemeManagerViewModel _viewModel = new();

    public ThemeManagerWindow()
    {
        InitializeComponent();
        DataContext = _viewModel;
    }

    private void OnCloseClick(object? sender, RoutedEventArgs e) => Close();

    /// <summary>Deleting a theme can't be undone and quietly moves any bot using it back to Default, so ask first.</summary>
    private async void OnDeleteClick(object? sender, RoutedEventArgs e)
    {
        if (_viewModel.SelectedTheme is not { IsBuiltIn: false } theme)
        {
            return;
        }

        var dialog = new Window
        {
            Title = "Delete Theme",
            Width = 420,
            SizeToContent = SizeToContent.Height,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
        };
        var delete = new Button { Content = "Delete", IsDefault = true };
        var cancel = new Button { Content = "Cancel", IsCancel = true };
        delete.Click += (_, _) => dialog.Close(true);
        cancel.Click += (_, _) => dialog.Close(false);
        dialog.Content = new StackPanel
        {
            Margin = new Thickness(20),
            Spacing = 12,
            Children =
            {
                new TextBlock { Text = $"Delete \"{theme.Name}\"?", FontWeight = FontWeight.Bold, TextWrapping = TextWrapping.Wrap },
                new TextBlock { Text = "Any bot using it goes back to the Default theme.", Opacity = 0.8, TextWrapping = TextWrapping.Wrap },
                new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Spacing = 8,
                    Children = { cancel, delete },
                },
            },
        };

        if (await dialog.ShowDialog<bool>(this))
        {
            _viewModel.DeleteCommand.Execute(null);
        }
    }
}
