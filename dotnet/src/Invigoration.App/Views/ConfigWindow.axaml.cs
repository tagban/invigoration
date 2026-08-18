using Avalonia.Controls;
using Avalonia.Interactivity;
using Invigoration.Core.Config;

namespace Invigoration.App.Views;

/// <summary>
/// Edits a BotConfig in place and returns it via ShowDialog on Save (null on
/// Cancel). Only used for creating a *new* bot right now — the config passed
/// in is freshly created and discarded if cancelled, so live-binding
/// mutation is safe. Editing an already-running bot's settings isn't wired
/// up yet (remove and re-add the bot tab as a workaround).
/// </summary>
public partial class ConfigWindow : Window
{
    private readonly BotConfig _config;

    public ConfigWindow() : this(new BotConfig())
    {
    }

    public ConfigWindow(BotConfig config)
    {
        InitializeComponent();
        _config = config;
        DataContext = _config;
    }

    private void OnSaveClick(object? sender, RoutedEventArgs e) => Close(_config);

    private void OnCancelClick(object? sender, RoutedEventArgs e) => Close(null);
}
