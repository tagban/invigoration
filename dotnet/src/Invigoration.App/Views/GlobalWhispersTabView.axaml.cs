using Avalonia.Controls;
using Avalonia.Input;
using Invigoration.App.ViewModels;

namespace Invigoration.App.Views;

public partial class GlobalWhispersTabView : UserControl
{
    public GlobalWhispersTabView()
    {
        InitializeComponent();

        var whisperInputBox = this.FindControl<TextBox>("WhisperInputBox");
        if (whisperInputBox is not null)
        {
            whisperInputBox.KeyDown += (_, e) =>
            {
                if (e.Key == Key.Enter && DataContext is GlobalWhispersTabViewModel { SelectedThread: { } thread })
                {
                    thread.Owner.SendWhisperCommand.Execute(thread);
                    e.Handled = true;
                }
            };
        }
    }
}
