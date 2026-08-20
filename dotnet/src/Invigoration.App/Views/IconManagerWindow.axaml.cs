using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Invigoration.App.ViewModels;

namespace Invigoration.App.Views;

public partial class IconManagerWindow : Window
{
    private readonly IconManagerViewModel _viewModel = new();

    public IconManagerWindow()
    {
        InitializeComponent();
        DataContext = _viewModel;
    }

    private async void OnChangeIconClick(object? sender, RoutedEventArgs e)
    {
        if (sender is not Button { DataContext: IconSlotViewModel slot })
        {
            return;
        }

        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = $"Choose an icon for {slot.DisplayName}",
            AllowMultiple = false,
            FileTypeFilter = [new FilePickerFileType("Images") { Patterns = ["*.png", "*.gif", "*.jpg", "*.jpeg", "*.bmp"] }],
        });

        var path = files.Count > 0 ? files[0].TryGetLocalPath() : null;
        if (path is null)
        {
            return;
        }

        _viewModel.ApplyIcon(slot, path);
    }

    private void OnCloseClick(object? sender, RoutedEventArgs e) => Close();
}
