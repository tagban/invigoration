using Avalonia.Controls;
using Avalonia.Input;

namespace Invigoration.App.Views;

public partial class HotlineTabView : UserControl
{
    public HotlineTabView() => InitializeComponent();

    private void OnTabHeaderContextRequested(object? sender, ContextRequestedEventArgs e) =>
        TabHeaderContextMenu.OnContextRequested(sender, e);
}
