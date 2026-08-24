using System.Collections.Specialized;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Layout;
using Avalonia.Threading;
using Invigoration.App.Models;
using Invigoration.App.ViewModels;

namespace Invigoration.App.Views;

/// <summary>Renders one ChannelTabViewModel's ChatLines — same hand-rolled append-to-Inlines approach as BotTabView.axaml.cs's flat log, just one instance per joined channel instead of one per bot.</summary>
public partial class ChannelTabView : UserControl
{
    private ChannelTabViewModel? _attachedViewModel;

    public ChannelTabView()
    {
        InitializeComponent();
        DataContextChanged += (_, _) => Attach();
    }

    private void Attach()
    {
        if (_attachedViewModel is not null)
        {
            _attachedViewModel.ChatLines.CollectionChanged -= OnChatLinesChanged;
        }

        _attachedViewModel = DataContext as ChannelTabViewModel;

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

        // Same fix as BotTabView.axaml.cs's AttachChatLog: repopulating the log above doesn't
        // move the ScrollViewer on its own, so switching to a different joined channel's
        // sub-tab would otherwise leave it wherever it last was instead of at the bottom.
        var scroll = this.FindControl<ScrollViewer>("ChatScroll");
        if (scroll is not null)
        {
            Dispatcher.UIThread.Post(() => scroll.ScrollToEnd(), DispatcherPriority.Background);
        }
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

        if (line.Icon is not null)
        {
            // Same Viewbox-wrapping fix as BotTabView.axaml.cs's AppendLine — a plain Image's own
            // Width/Height aren't reliably honored once embedded via InlineUIContainer.
            inlines.Add(new InlineUIContainer(new Viewbox
            {
                Width = 16,
                Height = 16,
                Stretch = Avalonia.Media.Stretch.Uniform,
                Margin = new Thickness(0, 0, 4, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new Image { Source = line.Icon },
            }));
        }

        foreach (var segment in line.Segments)
        {
            inlines.Add(new Run(segment.Text) { Foreground = segment.Brush });
        }
    }
}
