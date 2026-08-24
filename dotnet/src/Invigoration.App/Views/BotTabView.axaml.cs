using System.Collections.Specialized;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Controls.Primitives;
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

        // Resolved lazily inside the delegate (not here) since the Window ancestor may not
        // be attached yet at DataContextChanged time — by the time Connect actually fires,
        // it will be.
        _attachedViewModel.Engine.Sc2ChallengeHandler = (url, _) =>
        {
            var owner = this.FindAncestorOfType<Window>();
            return owner is null
                ? throw new InvalidOperationException("No window available to show the Battle.net login popup.")
                : Sc2LoginChallenge.ShowAsync(owner, url);
        };

        foreach (var line in _attachedViewModel.ChatLines)
        {
            AppendLine(chatText, line);
        }

        _attachedViewModel.ChatLines.CollectionChanged += OnChatLinesChanged;

        // Switching to this bot's tab repopulates the whole log above, but that alone doesn't
        // move the ScrollViewer — it keeps whatever offset it last had (from this bot's own
        // previous session, or a freshly-cleared 0/0 for one never viewed yet), not necessarily
        // the bottom. Scroll it to the end explicitly, same as a newly-arrived line does.
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

    /// <summary>
    /// A direct code-behind handler rather than a Command binding, deliberately — this button
    /// sits inside a TabControl.ItemTemplate nested two TabControls deep (the outer per-bot
    /// TabControl in MainWindow.axaml, then this bot's own per-channel one), and the
    /// $parent[TabControl].((vm:BotTabViewModel)DataContext) chain that would otherwise be
    /// needed to reach BotTabViewModel from there is exactly the kind of binding that fails
    /// silently (no exception, the button just never does anything) if it resolves to the
    /// wrong ancestor. Sender.DataContext is unambiguous: it's always this row's own
    /// ChannelTabViewModel, no tree-walking involved.
    /// </summary>
    /// <summary>
    /// The right-click "Whisper" ContextFlyout's Send button — sends through this bot's own
    /// engine (whichever product-appropriate path SendChatCommandAsync/SendSc2Async resolves
    /// to for a "/w " body) and closes the popup. Closing works by walking up to the hosting
    /// Popup and clearing IsOpen — Flyout content is always parented under one, so this is the
    /// same trick regardless of which friend row's flyout is open.
    /// </summary>
    private void OnSendInlineWhisperClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Control { DataContext: FriendEntryViewModel friend } control && DataContext is BotTabViewModel vm)
        {
            var text = friend.WhisperDraft.Trim();
            if (text.Length > 0)
            {
                friend.WhisperDraft = "";
                _ = vm.Engine.SendChatCommandAsync($"/w {friend.Account} {text}");
            }

            if (control.FindAncestorOfType<Popup>() is { } popup)
            {
                popup.IsOpen = false;
            }
        }
    }

    private void OnLeaveChannelClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Control { DataContext: ChannelTabViewModel tab } && DataContext is BotTabViewModel vm)
        {
            vm.LeaveChannelCommand.Execute(tab);
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
