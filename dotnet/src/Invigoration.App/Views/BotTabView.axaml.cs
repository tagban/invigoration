using System.Collections.Specialized;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
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

    private void OnFriendProfileInfoClick(object? sender, RoutedEventArgs e) => OpenProfileAndCloseFlyout(sender);

    private void OnFriendClanRankClick(object? sender, RoutedEventArgs e) => CloseFriendFlyoutAndRun(sender, friend => $"/claninfo {friend.Account}");

    private void OnFriendSquelchClick(object? sender, RoutedEventArgs e) => CloseFriendFlyoutAndRun(sender, friend => $"/squelch {friend.Account}");

    /// <summary>Populates the friend flyout's rank ComboBox straight from the global ClanRankStore on Loaded — no ancestor-DataContext walk needed since the store is a plain static list, sidestepping the "bindings through Popups are unreliable" issue noted elsewhere in this file.</summary>
    private void OnClanRankComboLoaded(object? sender, RoutedEventArgs e)
    {
        if (sender is ComboBox combo)
        {
            combo.ItemsSource = Invigoration.Core.Clan.ClanRankStore.Ranks.Select(r => r.Name).ToList();
        }
    }

    private void OnFriendSetClanRankClick(object? sender, RoutedEventArgs e)
    {
        if (sender is not Control { DataContext: Models.FriendEntryViewModel friend } control || DataContext is not BotTabViewModel vm)
        {
            return;
        }

        if (control.Parent is Panel panel && panel.Children.OfType<ComboBox>().FirstOrDefault() is { SelectedItem: string rank })
        {
            _ = vm.Engine.RunLocalCommandAsync($"/clanrank {friend.Account} {rank}");
        }

        if (control.FindAncestorOfType<Popup>() is { } popup)
        {
            popup.IsOpen = false;
        }
    }

    private void OpenProfileAndCloseFlyout(object? sender)
    {
        if (sender is Control { DataContext: Models.FriendEntryViewModel friend } control && DataContext is BotTabViewModel vm)
        {
            OpenProfileWindow(vm, friend.Account);
            if (control.FindAncestorOfType<Popup>() is { } popup)
            {
                popup.IsOpen = false;
            }
        }
    }

    private void CloseFriendFlyoutAndRun(object? sender, Func<Models.FriendEntryViewModel, string> buildCommand)
    {
        if (sender is Control { DataContext: Models.FriendEntryViewModel friend } control && DataContext is BotTabViewModel vm)
        {
            _ = vm.Engine.RunLocalCommandAsync(buildCommand(friend));
            if (control.FindAncestorOfType<Popup>() is { } popup)
            {
                popup.IsOpen = false;
            }
        }
    }

    /// <summary>Right-click menu on the classic Users list — MenuItem.Click, no manual popup-closing needed (ContextMenu closes itself on click, unlike the Friends list's ContextFlyout above).</summary>
    private void OnUserWhisperClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Control { DataContext: Models.ChannelUserViewModel user } && DataContext is BotTabViewModel vm
            && this.FindAncestorOfType<Window>()?.DataContext is MainWindowViewModel mainVm)
        {
            mainVm.FocusWhisperThread(vm, user.Username);
        }
    }

    private void OnUserProfileInfoClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Control { DataContext: Models.ChannelUserViewModel user } && DataContext is BotTabViewModel vm)
        {
            OpenProfileWindow(vm, user.Username);
        }
    }

    private void OpenProfileWindow(BotTabViewModel vm, string account)
    {
        var owner = this.FindAncestorOfType<Window>();
        var window = new ProfileWindow(vm.Engine, account);
        if (owner is not null)
        {
            _ = window.ShowDialog(owner);
        }
        else
        {
            window.Show();
        }
    }

    private void OnUserClanRankClick(object? sender, RoutedEventArgs e) => RunUserCommand(sender, user => $"/claninfo {user.Username}");

    private void OnUserSquelchClick(object? sender, RoutedEventArgs e) => RunUserCommand(sender, user => $"/squelch {user.Username}");

    private void OnUserAddFriendClick(object? sender, RoutedEventArgs e) => RunUserCommand(sender, user => $"/f add {user.Username}");

    /// <summary>
    /// Fired when the classic Users list's right-click menu opens — (re)populates the "Clan Rank"
    /// submenu's per-rank "Set: X" entries straight from ClanRankStore, after the two static items
    /// ("View Info" + a Separator) it starts with in XAML, and syncs "Classic Icon Style"'s
    /// checkmark to the bot's current Config.ClassicUserIconStyle (there's no reachable binding
    /// path from this row's own DataContext back up to Config — same reasoning as everywhere else
    /// in this file — so it's set here in code-behind instead, each time the menu opens).
    /// </summary>
    private void OnUserContextMenuOpened(object? sender, RoutedEventArgs e)
    {
        if (sender is not ContextMenu { PlacementTarget: Control { DataContext: Models.ChannelUserViewModel user } } menu
            || DataContext is not BotTabViewModel vm)
        {
            return;
        }

        if (menu.Items.OfType<MenuItem>().FirstOrDefault(m => m.Name == "UserClanRankMenu") is { } clanRankMenu)
        {
            while (clanRankMenu.Items.Count > 2)
            {
                clanRankMenu.Items.RemoveAt(clanRankMenu.Items.Count - 1);
            }

            foreach (var rank in Invigoration.Core.Clan.ClanRankStore.Ranks)
            {
                var item = new MenuItem { Header = $"Set: {rank.Name}" };
                item.Click += (_, _) => _ = vm.Engine.RunLocalCommandAsync($"/clanrank {user.Username} {rank.Name}");
                clanRankMenu.Items.Add(item);
            }
        }

        if (menu.Items.OfType<MenuItem>().FirstOrDefault(m => m.Name == "ClassicIconStyleMenuItem") is { } classicIconStyleItem)
        {
            classicIconStyleItem.IsChecked = vm.Config.ClassicUserIconStyle;
        }
    }

    private void OnToggleClassicIconStyleClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is BotTabViewModel vm)
        {
            vm.ToggleClassicUserIconStyle();
        }
    }

    private void RunUserCommand(object? sender, Func<Models.ChannelUserViewModel, string> buildCommand)
    {
        if (sender is Control { DataContext: Models.ChannelUserViewModel user } && DataContext is BotTabViewModel vm)
        {
            _ = vm.Engine.RunLocalCommandAsync(buildCommand(user));
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

        if (line.Icon is not null)
        {
            // A plain Image's own Width/Height aren't reliably honored once embedded via
            // InlineUIContainer inside a text flow — wrapping in a Viewbox forces the exact
            // rendered size regardless of how the surrounding line measures its inline content.
            // VerticalAlignment on that inner Viewbox does nothing for inline text-flow
            // positioning, though (confirmed live — the icon hung low, overlapping the line
            // below): text flow positions an InlineUIContainer via its own BaselineAlignment
            // property instead, which defaults to Top (the box's top pinned to the line's top,
            // so a too-tall icon spills downward past the baseline) — Center fixes that. But
            // BaselineAlignment alone wasn't enough either (confirmed live, still overlapping):
            // ChatText's LineHeight is now explicitly 20 (BotTabView.axaml), so a 20px icon was
            // still exactly as tall as the whole line with zero room to spare — any sub-pixel
            // rounding pushed it into the next line. 16px comfortably fits within that 20px line
            // box with margin either side, while still meeting "at least the height of the text"
            // (13px FontSize) that this size was originally bumped up from.
            inlines.Add(new InlineUIContainer(new Viewbox
            {
                Width = 16,
                Height = 16,
                Stretch = Avalonia.Media.Stretch.Uniform,
                Margin = new Thickness(0, 0, 4, 0),
                Child = new Image { Source = line.Icon },
            })
            {
                BaselineAlignment = Avalonia.Media.BaselineAlignment.Center,
            });
        }

        foreach (var segment in line.Segments)
        {
            inlines.Add(new Run(segment.Text) { Foreground = segment.Brush });
        }
    }
}
