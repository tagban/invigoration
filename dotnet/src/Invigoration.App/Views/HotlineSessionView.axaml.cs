using System.Collections.Specialized;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Invigoration.App.ViewModels;

namespace Invigoration.App.Views;

public partial class HotlineSessionView : UserControl
{
    private HotlineSessionViewModel? _attachedViewModel;

    public HotlineSessionView()
    {
        InitializeComponent();

        var inputBox = this.FindControl<TextBox>("InputBox");
        if (inputBox is not null)
        {
            inputBox.KeyDown += (_, e) =>
            {
                if (e.Key == Key.Enter && DataContext is HotlineSessionViewModel vm)
                {
                    vm.SendCommand.Execute(null);
                    e.Handled = true;
                }
            };
        }

        DataContextChanged += (_, _) => AttachAutoScroll();
    }

    /// <summary>
    /// Same fix as ChannelTabView.axaml.cs's own autoscroll — a plain ListBox never scrolls itself
    /// to a newly-added item, so the chat log otherwise stays wherever it last was as new lines
    /// keep arriving above the visible area. Fixed per direct user report ("AutoScroll is also not
    /// working on Hotline chats"). ScrollIntoView (not FindControl-ing the ListBox's own internal
    /// ScrollViewer and calling ScrollToEnd) since that's the documented Avalonia ListBox API for
    /// this and doesn't depend on reaching into the control's template.
    /// </summary>
    private void AttachAutoScroll()
    {
        if (_attachedViewModel is not null)
        {
            _attachedViewModel.Messages.CollectionChanged -= OnMessagesChanged;
        }

        _attachedViewModel = DataContext as HotlineSessionViewModel;
        if (_attachedViewModel is null)
        {
            return;
        }

        _attachedViewModel.Messages.CollectionChanged += OnMessagesChanged;
        ScrollToLastMessage();
    }

    private void OnMessagesChanged(object? sender, NotifyCollectionChangedEventArgs e) => ScrollToLastMessage();

    private void ScrollToLastMessage()
    {
        if (_attachedViewModel is not { Messages.Count: > 0 } vm)
        {
            return;
        }

        var listBox = this.FindControl<ListBox>("MessagesList");
        if (listBox is null)
        {
            return;
        }

        // Posted, not called inline — the ListBox needs a layout pass to realize the just-added
        // item's container before ScrollIntoView can actually find it.
        Dispatcher.UIThread.Post(() => listBox.ScrollIntoView(vm.Messages[^1]), DispatcherPriority.Background);
    }

    /// <summary>A ListBox only supports item selection, not text selection — this is the direct answer to "I can't copy from the chat log" rather than relying on click-drag text selection alone.</summary>
    private async void OnCopyLogClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not HotlineSessionViewModel vm)
        {
            return;
        }

        var clipboard = TopLevel.GetTopLevel(this)?.Clipboard;
        if (clipboard is not null)
        {
            await clipboard.SetTextAsync(string.Join(Environment.NewLine, vm.Messages.Select(m => m.FullText)));
        }
    }
}
