using Avalonia.Controls;
using Avalonia.Input;
using Invigoration.App.ViewModels;

namespace Invigoration.App.Views;

/// <summary>
/// The right-click menu on a tab's header: Connect/Disconnect on a single bot or Hotline server,
/// Connect All/Disconnect All/Reconnect All on a group (a bot tab group, or a Hotline tracker with
/// its servers). Built fresh each time it opens rather than bound, because the header templates
/// hold several unrelated view-model types and the menu has to reflect the state right now.
/// </summary>
internal static class TabHeaderContextMenu
{
    public static void OnContextRequested(object? sender, ContextRequestedEventArgs e)
    {
        if (sender is not Control header)
        {
            return;
        }

        var items = Build(header.DataContext);
        if (items.Count == 0)
        {
            return;
        }

        var menu = new ContextMenu();
        if (MenuTitle(header.DataContext) is { Length: > 0 } title)
        {
            menu.Items.Add(new MenuItem { Header = title, IsEnabled = false });
            menu.Items.Add(new Separator());
        }

        foreach (var item in items)
        {
            menu.Items.Add(item);
        }

        menu.Open(header);
        e.Handled = true;
    }

    /// <summary>Which tab the menu belongs to, so it's clear what "All" covers: a group says so.</summary>
    private static string? MenuTitle(object? tab) => tab switch
    {
        BotGroupTabViewModel group => $"Group: {group.Title}",
        HotlineTabViewModel hotline => $"Group: {hotline.Title}",
        BotTabViewModel bot => bot.Title,
        HotlineSessionViewModel session => session.Title,
        MusicTabViewModel => "Music",
        _ => null,
    };

    private static List<MenuItem> Build(object? tab) => tab switch
    {
        BotTabViewModel bot =>
        [
            Item(bot.ConnectButtonText, bot.CanConnect, () => bot.ConnectCommand.ExecuteAsync(null)),
            Item("Disconnect", bot.CanDisconnect, () => bot.DisconnectCommand.ExecuteAsync(null)),
        ],
        BotGroupTabViewModel group =>
        [
            Item("Connect All", group.AnyCanConnect, group.ConnectAllAsync),
            Item("Disconnect All", group.AnyCanDisconnect, group.DisconnectAllAsync),
            Item("Reconnect All", group.Bots.Count > 0, group.ReconnectAllAsync),
        ],
        HotlineSessionViewModel session =>
        [
            Item("Connect", session.CanConnectNow, session.ConnectNowAsync),
            Item("Disconnect", session.CanDisconnectNow, () => { session.GoOffline(); return Task.CompletedTask; }),
            Item("Reconnect", !session.CanConnectNow || session.IsReconnecting, session.ReconnectNowAsync),
            Item("Close", true, () => session.DisconnectCommand.ExecuteAsync(null)),
        ],
        HotlineTabViewModel hotline when hotline.HasSessions =>
        [
            Item("Connect All", hotline.AnyCanConnect, hotline.ConnectAllAsync),
            Item("Disconnect All", hotline.AnyCanDisconnect, () => { hotline.DisconnectAll(); return Task.CompletedTask; }),
            Item("Reconnect All", true, hotline.ReconnectAllAsync),
        ],
        // Spotify only has the one play/pause toggle, so this is worded from the tab's own playing state.
        MusicTabViewModel music =>
        [
            Item(music.IsPlaying ? "Stop" : "Start", true, () => music.PlayPauseCommand.ExecuteAsync(null)),
        ],
        _ => [],
    };

    private static MenuItem Item(string header, bool enabled, Func<Task> action)
    {
        var item = new MenuItem { Header = header, IsEnabled = enabled };
        item.Click += (_, _) => _ = action();
        return item;
    }
}
