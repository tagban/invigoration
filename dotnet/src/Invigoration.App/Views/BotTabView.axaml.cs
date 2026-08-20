using System.Collections.Specialized;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Invigoration.App.Models;
using Invigoration.App.ViewModels;

namespace Invigoration.App.Views;

public partial class BotTabView : UserControl
{
    private BotTabViewModel? _attachedViewModel;

    public BotTabView()
    {
        InitializeComponent();
        DataContextChanged += (_, _) => AttachChatLog();

        var inputBox = this.FindControl<TextBox>("InputBox");
        if (inputBox is not null)
        {
            inputBox.KeyDown += (_, e) =>
            {
                if (e.Key == Key.Enter && DataContext is BotTabViewModel vm)
                {
                    vm.SendCommand.Execute(null);
                    e.Handled = true;
                }
            };
        }
    }

    private void AttachChatLog()
    {
        if (_attachedViewModel is not null)
        {
            _attachedViewModel.ChatLines.CollectionChanged -= OnChatLinesChanged;
        }

        _attachedViewModel = DataContext as BotTabViewModel;

        var chatText = this.FindControl<SelectableTextBlock>("ChatText");
        chatText?.Inlines?.Clear();

        if (_attachedViewModel is null)
        {
            return;
        }

        foreach (var line in _attachedViewModel.ChatLines)
        {
            AppendLine(chatText, line);
        }

        _attachedViewModel.ChatLines.CollectionChanged += OnChatLinesChanged;
    }

    private void OnChatLinesChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        var chatText = this.FindControl<SelectableTextBlock>("ChatText");
        var scroll = this.FindControl<ScrollViewer>("ChatScroll");

        if (e.Action == NotifyCollectionChangedAction.Add && e.NewItems is not null)
        {
            foreach (ChatLineViewModel line in e.NewItems)
            {
                AppendLine(chatText, line);
            }
        }
        else if (e.Action == NotifyCollectionChangedAction.Reset)
        {
            chatText?.Inlines?.Clear();
        }

        if (scroll is not null)
        {
            Dispatcher.UIThread.Post(() => scroll.ScrollToEnd(), DispatcherPriority.Background);
        }
    }

    private async void OnManageClanClick(object? sender, RoutedEventArgs e)
    {
        var owner = this.FindAncestorOfType<Window>();
        if (owner is not null)
        {
            await new ClanWindow().ShowDialog(owner);
        }
    }

    private static void AppendLine(SelectableTextBlock? chatText, ChatLineViewModel line)
    {
        if (chatText?.Inlines is not { } inlines)
        {
            return;
        }

        if (inlines.Count > 0)
        {
            inlines.Add(new LineBreak());
        }

        foreach (var segment in line.Segments)
        {
            inlines.Add(new Run(segment.Text) { Foreground = segment.Brush });
        }
    }
}
