using Avalonia.Controls;
using Avalonia.Input;

namespace Invigoration.App.Views;

public partial class BotGroupTabView : UserControl
{
    public BotGroupTabView() => InitializeComponent();

    private void OnTabHeaderContextRequested(object? sender, ContextRequestedEventArgs e) =>
        TabHeaderContextMenu.OnContextRequested(sender, e);
}
