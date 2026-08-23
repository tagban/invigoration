using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Invigoration.App.ViewModels;

namespace Invigoration.App.Views;

public partial class ColorManagerWindow : Window
{
    private readonly ColorManagerViewModel _viewModel = new();

    public ColorManagerWindow()
    {
        InitializeComponent();
        DataContext = _viewModel;
    }

    private void OnCloseClick(object? sender, RoutedEventArgs e) => Close();

    private async void OnImportSchemeClick(object? sender, RoutedEventArgs e)
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Import Color Scheme",
            AllowMultiple = false,
            FileTypeFilter = [new FilePickerFileType("Invigoration Color Scheme") { Patterns = ["*.json"] }],
        });

        var path = files.Count > 0 ? files[0].TryGetLocalPath() : null;
        if (path is null)
        {
            return;
        }

        try
        {
            _viewModel.ImportSchemeFile(path);
        }
        catch (Exception ex) when (ex is IOException or System.Text.Json.JsonException)
        {
            // No dialog infrastructure exists yet to surface this; a bad file just silently doesn't import.
        }
    }
}
