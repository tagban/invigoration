using Avalonia.Controls;
using Avalonia.Interactivity;
using Invigoration.App.ViewModels;
using Invigoration.Core.Config;

namespace Invigoration.App.Views;

/// <summary>
/// Edits a clone of the given BotConfig and returns the edited clone via
/// ShowDialog on Save (null on Cancel) — the original is never mutated, so
/// this is safe for both adding a new bot and editing an already-added one;
/// the caller decides what to do with the result (use it as the new bot's
/// config, or assign it back onto an existing BotEngine.Config). Color
/// scheme and icon set editing both moved to their own windows under the
/// top-level Customize menu — this window only picks which saved one to use.
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
}
