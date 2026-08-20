using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Invigoration.App.ViewModels;
using Invigoration.Core.Config;

namespace Invigoration.App.Views;

/// <summary>
/// Edits a clone of the given BotConfig and returns the edited clone via
/// ShowDialog on Save (null on Cancel) — the original is never mutated, so
/// this is safe for both adding a new bot and editing an already-added one;
/// the caller decides what to do with the result (use it as the new bot's
/// config, or assign it back onto an existing BotEngine.Config).
/// </summary>
public partial class ConfigWindow : Window
{
    private readonly ConfigViewModel _viewModel;

    public ConfigWindow() : this(new BotConfig())
    {
    }

    public ConfigWindow(BotConfig config)
    {
        InitializeComponent();
        _viewModel = new ConfigViewModel(BotConfig.Clone(config));
        DataContext = _viewModel;
    }

    private void OnSaveClick(object? sender, RoutedEventArgs e) => Close(_viewModel.Config);

    private void OnCancelClick(object? sender, RoutedEventArgs e) => Close(null);

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
