using System.Diagnostics;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Controls.Primitives;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Invigoration.App.ViewModels;
using Invigoration.Core;
using Invigoration.Core.Config;
using Invigoration.Core.Music;

namespace Invigoration.App.Views;

public partial class MainWindow : Window
{
    private const string BaseTitle = $"Invigoration v{AppVersion.Current}";

    private static readonly TimeSpan TitleUpdateInterval = TimeSpan.FromSeconds(15);

    /// <summary>Kept so the static PrivateMessageRequested subscription can be removed when this window closes.</summary>
    private Action<Models.WhisperThreadViewModel>? _focusHotlineWhisper;

    public MainWindow()
    {
        InitializeComponent();
        Title = BaseTitle;
        Closing += (_, _) =>
        {
            ViewModel?.SaveAll();
            Invigoration.Core.Clan.ClanRosterStore.FlushPendingSave();
            if (_focusHotlineWhisper is not null)
            {
                HotlineUserRowViewModel.PrivateMessageRequested -= _focusHotlineWhisper;
            }
        };
        StartTitleUpdateTimer();
        // TopLevelTabs puts the Whispers pseudo-tab first, then (if enabled) Music — default to
        // the first real bot/group instead, matching this window's behavior before either
        // existed. A hardcoded SelectedIndex=1 (this used to be that) would land on Music instead
        // of a bot whenever Music is enabled, since it now also occupies an early slot. Loaded
        // (not the constructor) since DataContext isn't set yet at construction time.
        Loaded += (_, _) =>
        {
            if (this.FindControl<TabStrip>("TopLevelTabControl") is { } tabControl && ViewModel is { } vm)
            {
                var firstBotTab = vm.TopLevelTabs.FirstOrDefault(t => t is BotTabViewModel or BotGroupTabViewModel);
                if (firstBotTab is not null)
                {
                    tabControl.SelectedItem = firstBotTab;
                }

                vm.PropertyChanged += OnViewModelPropertyChanged;
                RebuildIconSetMenus();

                // A Hotline "Send Private Message..." opens the thread and asks to be shown.
                // Selecting it is all that's needed — OnViewModelPropertyChanged already switches
                // the strip to Whispers whenever the selected thread changes, the same path the
                // bot-side right-click uses. Unsubscribed on close: the event is static, so a
                // handler holding this window would outlive it.
                _focusHotlineWhisper = vm.FocusWhisperThread;
                HotlineUserRowViewModel.PrivateMessageRequested += _focusHotlineWhisper;

                // One-time questions, one after the other, posted so the window is up first: whether
                // to use Spotify at all, then — for a bot already on a character-dock theme from
                // before that offer existed — the D2 gear data.
                Dispatcher.UIThread.Post(async () =>
                {
                    await AskAboutSpotifyAsync();
                    if (vm.Bots.Any(b => b.Theme.UsesCharacterDock))
                    {
                        await OfferD2EquipmentDownloadAsync();
                    }

                    // Last, and asks nothing: it either puts one dismissible line above the tabs or
                    // does nothing at all. Never blocks startup — a slow or unreachable GitHub just
                    // means no notice this run.
                    await vm.CheckForUpdatesAsync();
                });
            }
        };
    }

    /// <summary>
    /// Shows what's currently playing in the title bar (e.g. "Invigoration v2.0.3b - Spotify: Pink
    /// Pony Club") whenever Spotify is connected and something's playing (not paused), falling back
    /// to the plain BaseTitle otherwise. Polling (not event-driven) since GetNowPlayingAsync is
    /// pull-based — nothing raises an event when the track changes; DispatcherTimer (not a
    /// background Task/PeriodicTimer) since this only needs to run while the window exists and
    /// writes directly to a UI property.
    /// </summary>
    private void StartTitleUpdateTimer()
    {
        var timer = new DispatcherTimer { Interval = TitleUpdateInterval };
        timer.Tick += async (_, _) => await UpdateTitleAsync();
        timer.Start();
    }

    private async Task UpdateTitleAsync()
    {
        if (MusicPlayerRegistry.Controller is not { } controller)
        {
            Title = BaseTitle;
        }
        else
        {
            var nowPlaying = await controller.GetNowPlayingAsync();
            Title = nowPlaying is { IsPlaying: true } ? $"{BaseTitle} - {nowPlaying.Service}: {nowPlaying.Title}" : BaseTitle;
        }

        // Same tick also refreshes the optional bottom playback bar — no separate poll loop
        // needed for it (see MusicBarViewModel's remarks).
        if (ViewModel is { IsMusicBarEnabled: true } vm)
        {
            await vm.MusicBar.RefreshAsync();
        }
    }

    /// <summary>
    /// Toggling the Customize menu's "Music Player" checkbox rebuilds TopLevelTabs (RefreshTopLevelTabs), which — same reset-to-first-item issue AddBot/EditBot already had to work around via SelectTopLevelBot — would otherwise leave the tab strip showing Whispers instead of the Music tab the user just turned on.
    /// Also mirrors MainWindowViewModel.FocusWhisperThread (the right-click "Whisper" action) onto the actual TabStrip control — SelectedGlobalWhisperThread itself can't drive tab selection directly since TopLevelTabs' selection isn't bound TwoWay (see SelectedTopLevelItem's remarks).
    /// </summary>
    private void OnViewModelPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(MainWindowViewModel.SelectedBotIconSetName) or nameof(MainWindowViewModel.IconSetNames))
        {
            RebuildIconSetMenus();
        }

        if (ViewModel is not { } vm || this.FindControl<TabStrip>("TopLevelTabControl") is not { } tabControl)
        {
            return;
        }

        if (e.PropertyName == nameof(MainWindowViewModel.IsMusicEnabled) && vm.IsMusicEnabled)
        {
            tabControl.SelectedItem = vm.MusicTab;
        }
        else if (e.PropertyName == nameof(MainWindowViewModel.SelectedGlobalWhisperThread) && vm.SelectedGlobalWhisperThread is not null)
        {
            tabControl.SelectedItem = vm.WhispersTab;
        }
    }

    private MainWindowViewModel? ViewModel => DataContext as MainWindowViewModel;

    /// <summary>
    /// Bot → Icon Set, in both menus: every set, the selected bot's ticked, then Manage Icons.
    /// Built here rather than in XAML because saved sets come and go, and rebuilt whenever the list
    /// or the tick changes; the macOS menu never moves a tick by itself.
    /// </summary>
    private void RebuildIconSetMenus()
    {
        if (ViewModel is not { } vm)
        {
            return;
        }

        var names = vm.IconSetNames;
        var current = vm.SelectedBotIconSetName;
        var built = string.Join("\n", names.Prepend(current));
        if (built == _iconSetMenusBuiltFor)
        {
            return;
        }

        _iconSetMenusBuiltFor = built;

        if (FindNativeMenuItem(NativeMenu.GetMenu(this), "Icon Set")?.Menu is { } nativeMenu)
        {
            nativeMenu.Items.Clear();
            foreach (var name in names)
            {
                var item = new NativeMenuItem { Header = name.Replace("_", "__"), ToggleType = MenuItemToggleType.Radio, IsChecked = name == current };
                item.Click += (_, _) => UseIconSetForSelectedBot(name);
                nativeMenu.Items.Add(item);
            }

            nativeMenu.Items.Add(new NativeMenuItemSeparator());
            var manage = new NativeMenuItem { Header = "Manage Icons..." };
            manage.Click += OnManageIconsNativeClick;
            nativeMenu.Items.Add(manage);
        }

        IconSetMenu.Items.Clear();
        foreach (var name in names)
        {
            var item = new MenuItem { Header = name.Replace("_", "__"), ToggleType = MenuItemToggleType.Radio, GroupName = "IconSet", IsChecked = name == current };
            item.Click += (_, _) => UseIconSetForSelectedBot(name);
            IconSetMenu.Items.Add(item);
        }

        IconSetMenu.Items.Add(new Separator());
        var manageItem = new MenuItem { Header = "Manage _Icons..." };
        manageItem.Click += OnManageIconsClick;
        IconSetMenu.Items.Add(manageItem);
    }

    /// <summary>The list and tick the Icon Set menus were last built for, so the frequent "look again" from the Bot menu only rebuilds them when something they show has changed.</summary>
    private string? _iconSetMenusBuiltFor;

    private void UseIconSetForSelectedBot(string name)
    {
        if (ViewModel is { SelectedBot: { } bot } vm)
        {
            vm.UseIconSet(bot, name);
        }
    }

    private static NativeMenuItem? FindNativeMenuItem(NativeMenu? menu, string header)
    {
        foreach (var item in menu?.Items.OfType<NativeMenuItem>() ?? [])
        {
            if (item.Header == header)
            {
                return item;
            }

            if (FindNativeMenuItem(item.Menu, header) is { } found)
            {
                return found;
            }
        }

        return null;
    }

    /// <summary>Keeps MainWindowViewModel.SelectedBot in sync with whichever tab is actually showing — TabControl.SelectedItem can't bind directly to it two-way any more since TopLevelTabs is a mixed BotTabViewModel/GlobalWhispersTabViewModel collection; selecting the Whispers tab leaves SelectedBot as whatever bot was last actually selected, which is a reasonable "last bot you were looking at" fallback for the Bot menu (which names that bot at the top so it's never a guess). Also feeds SetActiveTopLevelItem, which drives every bot's IsActive/HasUnread state (see RecomputeActiveBot).</summary>
    private void OnTabHeaderContextRequested(object? sender, ContextRequestedEventArgs e) =>
        TabHeaderContextMenu.OnContextRequested(sender, e);

    private void OnTopLevelTabSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (ViewModel is not { } vm)
        {
            return;
        }

        var selected = e.AddedItems.Count > 0 ? e.AddedItems[0] : null;
        vm.SetActiveTopLevelItem(selected);
        if (selected is BotTabViewModel bot)
        {
            vm.SelectedBot = bot;
        }
    }

    // Each action below has two Click handlers with identical bodies: one typed for the
    // in-window Menu (RoutedEventArgs) and one for the macOS NativeMenu (plain EventArgs).
    // Avalonia's XAML compiler requires an exact delegate match, so the two can't share a
    // single method despite RoutedEventArgs being an EventArgs.

    private async void OnAddBotClick(object? sender, RoutedEventArgs e) => await AddBot();
    private async void OnAddBotNativeClick(object? sender, EventArgs e) => await AddBot();

    private async Task AddBot()
    {
        var dialog = new ConfigWindow(new BotConfig());
        var result = await dialog.ShowDialog<BotConfig?>(this);
        if (result is not null && ViewModel is { } vm)
        {
            vm.AddBot(result);
            if (ThemeLibrary.ResolveFor(result).Layout == ThemeLayout.CharacterDock)
            {
                await OfferD2EquipmentDownloadAsync();
            }

            // AddBot's own RefreshTopLevelTabs() call clears and repopulates TopLevelTabs, which
            // resets the TabControl's own selection to index 0 (the Whispers pseudo-tab, always
            // first) — SelectedBot is a separate ViewModel property, not something the TabControl
            // actually reads. Restore it to the bot AddBot just set SelectedBot to.
            if (vm.SelectedBot is { } added)
            {
                SelectTopLevelBot(added);
            }
        }
    }

    private void OnAddHotlineTrackerClick(object? sender, RoutedEventArgs e) => AddHotlineTracker();
    private void OnAddHotlineTrackerNativeClick(object? sender, EventArgs e) => AddHotlineTracker();

    private void AddHotlineTracker()
    {
        if (ViewModel is not { } vm)
        {
            return;
        }

        var tracker = vm.AddHotlineTracker();
        // Same TabControl-selection-reset issue as AddBot — RefreshTopLevelTabs() rebuilding
        // TopLevelTabs from scratch otherwise leaves the tab strip showing Whispers instead of
        // the tracker that was just added.
        if (this.FindControl<TabStrip>("TopLevelTabControl") is { } tabControl)
        {
            tabControl.SelectedItem = tracker;
        }
    }

    private async void OnEditBotClick(object? sender, RoutedEventArgs e) => await EditSelectedBot();
    private async void OnEditBotNativeClick(object? sender, EventArgs e) => await EditSelectedBot();

    private async Task EditSelectedBot()
    {
        var vm = ViewModel;
        if (vm?.SelectedBot is not { } selected)
        {
            return;
        }

        await ApplyBotEditAsync(selected, await new ConfigWindow(selected.Config).ShowDialog<BotConfig?>(this));
    }

    private async void OnBotAppearanceClick(object? sender, RoutedEventArgs e) => await EditSelectedBotAppearance();
    private async void OnBotAppearanceNativeClick(object? sender, EventArgs e) => await EditSelectedBotAppearance();

    /// <summary>Bot → Appearance: theme, colors and icons, applied the same way as a settings edit.</summary>
    private async Task EditSelectedBotAppearance()
    {
        var vm = ViewModel;
        if (vm?.SelectedBot is not { } selected)
        {
            await InformAsync("Bot Appearance", "Select a bot first", "Open the tab of the bot whose appearance you want to change, then choose Bot → Appearance.");
            return;
        }

        await ApplyBotEditAsync(selected, await new BotAppearanceWindow(selected.Config).ShowDialog<BotConfig?>(this));
    }

    /// <summary>Applies an edited copy of a bot's config from either window, then saves and keeps that bot's tab selected.</summary>
    private async Task ApplyBotEditAsync(BotTabViewModel selected, BotConfig? result)
    {
        var vm = ViewModel;
        if (vm is null)
        {
            return;
        }

        if (result is not null)
        {
            selected.ApplyConfig(result);
            Models.IconSets.ApplyIfNotShowing(selected.Config.IconSetName);
            vm.RefreshTopLevelTabs();
            vm.SaveAll();
            if (selected.Theme.UsesCharacterDock)
            {
                await OfferD2EquipmentDownloadAsync();
            }

            // Same TabControl-selection-reset issue as AddBot — RefreshTopLevelTabs() rebuilding
            // TopLevelTabs from scratch otherwise leaves the tab strip showing Whispers instead of
            // the bot that was actually just edited.
            SelectTopLevelBot(selected);
        }
    }

    /// <summary>Selects a bot's own top-level tab directly if it's ungrouped, or its containing group's tab (and that bot within the group's own nested TabControl) if it's now grouped — used after RefreshTopLevelTabs() rebuilds TopLevelTabs and resets the TabControl's own selection.</summary>
    private void SelectTopLevelBot(BotTabViewModel bot)
    {
        if (this.FindControl<TabStrip>("TopLevelTabControl") is not { } tabControl || ViewModel is not { } vm)
        {
            return;
        }

        if (vm.TopLevelTabs.Contains(bot))
        {
            tabControl.SelectedItem = bot;
            return;
        }

        var group = vm.TopLevelTabs.OfType<BotGroupTabViewModel>().FirstOrDefault(g => g.Bots.Contains(bot));
        if (group is null)
        {
            return;
        }

        group.SelectedBot = bot;
        tabControl.SelectedItem = group;
    }

    private void OnOpenConfigFolderClick(object? sender, RoutedEventArgs e) => OpenConfigFolder();
    private void OnOpenConfigFolderNativeClick(object? sender, EventArgs e) => OpenConfigFolder();

    private static void OpenConfigFolder()
    {
        var dir = ConfigStore.DefaultConfigDirectory();
        Directory.CreateDirectory(dir);
        Process.Start(new ProcessStartInfo(dir) { UseShellExecute = true });
    }

    private async void OnNormalizePasswordClick(object? sender, RoutedEventArgs e) => await NormalizeSelectedBotPassword();
    private async void OnNormalizePasswordNativeClick(object? sender, EventArgs e) => await NormalizeSelectedBotPassword();

    /// <summary>
    /// "Normalize Password to Battle.net Spec" (see BotEngine.NormalizePasswordAsync): confirms
    /// first, since it really does change the account's password on the server. When it can't
    /// apply (wrong logon system, no password, already lowercase) the engine explains why in that
    /// bot's own chat log instead, without asking anything.
    /// </summary>
    private async Task NormalizeSelectedBotPassword()
    {
        if (ViewModel?.SelectedBot is not { } selected)
        {
            return;
        }

        var engine = selected.Engine;
        if (engine.NormalizePasswordUnavailableReason is null)
        {
            var server = string.IsNullOrWhiteSpace(selected.Config.BattlenetServer) ? "the server" : selected.Config.BattlenetServer;
            var confirmed = await ConfirmAsync(
                "Normalize Password",
                $"Change the password for \"{selected.Config.Username}\" on {server} to all lowercase?",
                "Blizzard's game clients always send a password in lowercase, but Invigoration used to send it exactly as " +
                "typed — so if the password on this account was set with capital letters by the bot, a real game client can't log into it.\n\n" +
                "This reconnects the bot, changes the account's password on Battle.net from the capitalized version to the " +
                "lowercase one, and saves the lowercase version as this bot's password.",
                "Change Password");
            if (!confirmed)
            {
                return;
            }
        }

        await engine.NormalizePasswordAsync();
    }

    /// <summary>A small modal yes/no prompt — heading, explanation, and a confirm button beside Cancel.</summary>
    /// <summary>
    /// First run: controlling Spotify needs Premium, so rather than show a Music tab most people
    /// can't use, ask once. Yes keeps the tab and opens it; No hides the tab and the player bar
    /// (Customize → Music Player brings it back). Closing the question without answering asks
    /// again next launch. Never asked once Spotify is connected.
    /// </summary>
    private async Task AskAboutSpotifyAsync()
    {
        if (MusicSettingsStore.SpotifyPromptAnswered || ViewModel is not { } vm)
        {
            return;
        }

        if (MusicSettingsStore.SpotifyRefreshToken.Length > 0)
        {
            MusicSettingsStore.SpotifyPromptAnswered = true;
            return;
        }

        var answer = await AskAsync(
            "Spotify",
            "Do you have Spotify Premium and want to use it?",
            "Invigoration can show what's playing on your Spotify and let your bots skip, pause and save tracks from chat — " +
            "so you can whisper your bot to skip a song while you're in a game. Controlling Spotify needs a Premium account.\n\n" +
            "Music plays through Spotify itself, so you'll need either the Spotify app installed on this computer, or Spotify open and signed in in your web browser.\n\n" +
            "If you choose No, the Music tab stays hidden — you can turn it on any time from Customize → Music Player.",
            "Yes, use Spotify",
            "No",
            SpotifyAppDownload());
        if (answer is not { } useSpotify)
        {
            return;
        }

        MusicSettingsStore.SpotifyPromptAnswered = true;
        if (useSpotify)
        {
            vm.IsMusicEnabled = true;
            if (this.FindControl<TabStrip>("TopLevelTabControl") is { } tabControl)
            {
                tabControl.SelectedItem = vm.MusicTab;
            }
        }
        else
        {
            vm.IsMusicEnabled = false;
            vm.IsMusicBarEnabled = false;
        }
    }

    /// <summary>Spotify's download page for this computer's OS, as a link for the Spotify question.</summary>
    private static (string Text, string Url) SpotifyAppDownload() =>
        OperatingSystem.IsMacOS() ? ("Get the Spotify app for Mac", "https://www.spotify.com/download/mac/")
        : OperatingSystem.IsWindows() ? ("Get the Spotify app for Windows", "https://www.spotify.com/download/windows/")
        : ("Get the Spotify app for Linux", "https://www.spotify.com/download/linux/");

    private async Task<bool> ConfirmAsync(string title, string heading, string body, string confirmText) =>
        await AskAsync(title, heading, body, confirmText, "Cancel") == true;

    /// <summary>A two-button question, optionally with a link that opens in the browser without answering. True for the confirm button, false for the other, null when the window was closed without either.</summary>
    private async Task<bool?> AskAsync(string title, string heading, string body, string confirmText, string cancelText, (string Text, string Url)? link = null)
    {
        var dialog = new Window
        {
            Title = title,
            Width = 460,
            SizeToContent = SizeToContent.Height,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
        };

        var confirm = new Button { Content = confirmText, IsDefault = true };
        var cancel = new Button { Content = cancelText, IsCancel = true };
        confirm.Click += (_, _) => dialog.Close((bool?)true);
        cancel.Click += (_, _) => dialog.Close((bool?)false);

        var content = new StackPanel
        {
            Margin = new Avalonia.Thickness(20),
            Spacing = 12,
            Children =
            {
                new TextBlock { Text = heading, FontWeight = Avalonia.Media.FontWeight.Bold, TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                new TextBlock { Text = body, TextWrapping = Avalonia.Media.TextWrapping.Wrap, Opacity = 0.8 },
            },
        };

        if (link is { } l)
        {
            var linkButton = new Button
            {
                Content = l.Text,
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Left,
                Background = Avalonia.Media.Brushes.Transparent,
                Foreground = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.FromRgb(0x1E, 0xD7, 0x60)),
                Padding = new Avalonia.Thickness(0),
            };
            linkButton.Click += (_, _) => Process.Start(new ProcessStartInfo(l.Url) { UseShellExecute = true });
            content.Children.Add(linkButton);
        }

        content.Children.Add(new StackPanel
        {
            Orientation = Avalonia.Layout.Orientation.Horizontal,
            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right,
            Spacing = 8,
            Children = { cancel, confirm },
        });
        dialog.Content = content;

        return await dialog.ShowDialog<bool?>(this);
    }

    private async void OnRemoveBotClick(object? sender, RoutedEventArgs e) => await RemoveSelectedBot();
    private async void OnRemoveBotNativeClick(object? sender, EventArgs e) => await RemoveSelectedBot();

    /// <summary>
    /// Asks first: it deletes the bot's settings for good, and from a tab that isn't a bot the
    /// "selected" bot is only the last one that was showing — the question names it, so the wrong
    /// bot can't go quietly.
    /// </summary>
    private async Task RemoveSelectedBot()
    {
        var vm = ViewModel;
        if (vm?.SelectedBot is not { } selected)
        {
            return;
        }

        var confirmed = await ConfirmAsync(
            "Remove Bot",
            $"Remove \"{selected.Title}\"?",
            "Its settings are deleted from Invigoration and it disconnects if it's online. There's no undo — it would have to be added again from scratch.",
            "Remove");
        if (confirmed)
        {
            vm.RemoveBot(selected);
        }
    }

    private async void OnAboutClick(object? sender, RoutedEventArgs e) => await ShowAbout();
    private async void OnAboutNativeClick(object? sender, EventArgs e) => await ShowAbout();

    private async Task ShowAbout() => await new AboutWindow().ShowDialog(this);

    private async void OnClanMembersClick(object? sender, RoutedEventArgs e) => await ShowClanMembers();
    private async void OnClanMembersNativeClick(object? sender, EventArgs e) => await ShowClanMembers();

    private async Task ShowClanMembers() => await new ClanWindow().ShowDialog(this);

    private async void OnClanRanksClick(object? sender, RoutedEventArgs e) => await ShowClanRanks();
    private async void OnClanRanksNativeClick(object? sender, EventArgs e) => await ShowClanRanks();

    private async Task ShowClanRanks() => await new ClanRanksWindow().ShowDialog(this);

    private async void OnManageIconsClick(object? sender, RoutedEventArgs e) => await ShowIconManager();
    private async void OnManageIconsNativeClick(object? sender, EventArgs e) => await ShowIconManager();

    private async Task ShowIconManager() => await new IconManagerWindow().ShowDialog(this);

    private async void OnManageColorsClick(object? sender, RoutedEventArgs e) => await ShowColorManager();
    private async void OnManageColorsNativeClick(object? sender, EventArgs e) => await ShowColorManager();

    private async Task ShowColorManager() => await new ColorManagerWindow().ShowDialog(this);

    private async void OnDownloadD2EquipmentClick(object? sender, RoutedEventArgs e) => await DownloadD2EquipmentAsync(askFirst: true);
    private async void OnDownloadD2EquipmentNativeClick(object? sender, EventArgs e) => await DownloadD2EquipmentAsync(askFirst: true);

    private const string D2EquipmentHeading = "Show what Diablo II characters are wearing?";

    private static string D2EquipmentBody =>
        "Diablo II characters in the bottom strip can show what they're wearing when you hover over them — helm, armor, weapons " +
        $"and shield. That needs Diablo II gear data downloaded from {D2EquipmentStore.DefaultHost} and kept on this computer for every " +
        $"server; a bot connected to {D2EquipmentStore.DefaultHost} later picks up newer copies on its own. Nothing about you or your bots is sent.";

    /// <summary>
    /// The one-time automatic offer, made when a bot is (or turns out to be) on a character-dock
    /// theme: only if no gear list is stored yet and it hasn't been turned down before. "No"
    /// is remembered; the Customize menu item still works any time.
    /// </summary>
    private async Task OfferD2EquipmentDownloadAsync()
    {
        if (D2EquipmentStore.Current is not null || D2EquipmentStore.Declined || _offeringD2Equipment)
        {
            return;
        }

        _offeringD2Equipment = true;
        try
        {
            if (!await ConfirmAsync("Diablo II Gear Data", D2EquipmentHeading, D2EquipmentBody, "Download"))
            {
                D2EquipmentStore.Declined = true;
                return;
            }

            await DownloadD2EquipmentAsync(askFirst: false);
        }
        finally
        {
            _offeringD2Equipment = false;
        }
    }

    private bool _offeringD2Equipment;

    private async Task DownloadD2EquipmentAsync(bool askFirst)
    {
        if (askFirst && !await ConfirmAsync("Diablo II Gear Data", D2EquipmentHeading, D2EquipmentBody, "Download"))
        {
            return;
        }

        var (result, detail) = await D2EquipmentStore.DownloadAsync(D2EquipmentStore.TrustedServers);
        var (heading, body) = result switch
        {
            D2EquipmentDownloadResult.Saved => ("Diablo II gear data downloaded.",
                "Hover over a Diablo II character in the bottom strip to see what they're wearing. Characters on a realm that " +
                "doesn't keep items (every slot empty) won't show anything."),
            D2EquipmentDownloadResult.NotOnServer => ("The gear data isn't available yet.",
                $"{D2EquipmentStore.DefaultHost} isn't serving it yet. Try again later from Customize → Download Diablo II Gear Data."),
            D2EquipmentDownloadResult.Unreadable => ("The gear data couldn't be used.",
                $"The server sent a version this Invigoration doesn't understand ({detail}). An update to Invigoration may be needed."),
            _ => ("The download didn't finish.", $"Couldn't get the gear data ({detail}). Try again later."),
        };
        await InformAsync("Diablo II Gear Data", heading, body);
    }

    /// <summary>A small modal notice with a single OK.</summary>
    private async Task InformAsync(string title, string heading, string body)
    {
        var dialog = new Window
        {
            Title = title,
            Width = 440,
            SizeToContent = SizeToContent.Height,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
        };
        var ok = new Button { Content = "OK", IsDefault = true, IsCancel = true, HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right };
        ok.Click += (_, _) => dialog.Close();
        dialog.Content = new StackPanel
        {
            Margin = new Avalonia.Thickness(20),
            Spacing = 12,
            Children =
            {
                new TextBlock { Text = heading, FontWeight = Avalonia.Media.FontWeight.Bold, TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                new TextBlock { Text = body, TextWrapping = Avalonia.Media.TextWrapping.Wrap, Opacity = 0.8 },
                ok,
            },
        };
        await dialog.ShowDialog(this);
    }

    private async void OnManageThemesClick(object? sender, RoutedEventArgs e) => await ShowThemeManager();
    private async void OnManageThemesNativeClick(object? sender, EventArgs e) => await ShowThemeManager();

    private async Task ShowThemeManager() => await new ThemeManagerWindow().ShowDialog(this);

    private async void OnManageBattlenetProfilesClick(object? sender, RoutedEventArgs e) => await ShowBattlenetProfiles();
    private async void OnManageBattlenetProfilesNativeClick(object? sender, EventArgs e) => await ShowBattlenetProfiles();

    private async Task ShowBattlenetProfiles() => await new BattlenetCredentialProfilesWindow().ShowDialog(this);

    private void OnExitClick(object? sender, RoutedEventArgs e) => ExitApplication();
    private void OnExitNativeClick(object? sender, EventArgs e) => ExitApplication();

    private static void ExitApplication()
    {
        if (Avalonia.Application.Current?.ApplicationLifetime is
            Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop)
        {
            desktop.Shutdown();
        }
    }
}
