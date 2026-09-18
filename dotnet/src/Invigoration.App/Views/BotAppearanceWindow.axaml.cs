using Avalonia.Controls;
using Avalonia.Interactivity;
using Invigoration.App.ViewModels;
using Invigoration.Core.Config;

namespace Invigoration.App.Views;

/// <summary>
/// Bot → Appearance: one bot's theme, chat colors, icon set and tab group — everything
/// about how a bot looks, kept out of the bot settings so those stay about connecting. Same
/// contract as ConfigWindow: edits a clone and returns it on Save (null on Cancel), so the caller
/// applies it exactly as it applies a settings edit.
/// </summary>
public partial class BotAppearanceWindow : Window
{
    private readonly ConfigViewModel _viewModel;

    public BotAppearanceWindow() : this(new BotConfig())
    {
    }

    public BotAppearanceWindow(BotConfig config)
    {
        InitializeComponent();
        _viewModel = new ConfigViewModel(BotConfig.Clone(config));
        DataContext = _viewModel;
    }

    private void OnSaveClick(object? sender, RoutedEventArgs e) => Close(_viewModel.Config);

    private void OnCancelClick(object? sender, RoutedEventArgs e) => Close(null);
}
