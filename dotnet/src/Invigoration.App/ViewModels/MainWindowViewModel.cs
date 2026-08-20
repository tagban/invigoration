using System.Collections.ObjectModel;
using System.Diagnostics;
using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.Core;
using Invigoration.Core.Config;
using Invigoration.Core.Trivia;

namespace Invigoration.App.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    private readonly ConfigStore _store = new();

    public ObservableCollection<BotTabViewModel> Bots { get; } = [];

    [ObservableProperty]
    public partial BotTabViewModel? SelectedBot { get; set; }

    /// <summary>
    /// True once at least one configured bot has ClanFeatureEnabled on — the
    /// top-level "Clan" menu binds its IsVisible to this so it disappears
    /// entirely rather than offering clan management nobody's turned on.
    /// Recomputed wherever bot config can change (load, add, remove, and
    /// SaveAll after an edit) since BotConfig itself isn't observable.
    /// </summary>
    [ObservableProperty]
    public partial bool AnyBotHasClanEnabled { get; set; }

    public MainWindowViewModel()
    {
        foreach (var config in _store.Load())
        {
            Bots.Add(new BotTabViewModel(new BotEngine(config)));
        }

        SelectedBot = Bots.Count > 0 ? Bots[0] : null;
        RefreshAnyBotHasClanEnabled();
        _ = AutoConnectStartupBotsAsync();
        // Fire-and-forget and best-effort: seeds the Trivia folder with the base packs from
        // GitHub the first time (see TriviaPackDownloader), never blocks startup, and is a
        // no-op on every later launch once those files already exist locally.
        _ = TriviaPackDownloader.EnsureDownloadedAsync(err => Debug.WriteLine(err));
    }

    private void RefreshAnyBotHasClanEnabled() => AnyBotHasClanEnabled = Bots.Any(b => b.Config.ClanFeatureEnabled);

    /// <summary>
    /// Connects every bot with AutoConnectOnStartup, staggered a couple
    /// seconds apart rather than all at once — several bots opening TCP
    /// connections to the same server in the same instant is exactly the
    /// kind of burst a per-IP connection or flood limit can catch.
    /// </summary>
    private async Task AutoConnectStartupBotsAsync()
    {
        foreach (var bot in Bots.Where(b => b.Config.AutoConnectOnStartup).ToList())
        {
            await bot.ConnectCommand.ExecuteAsync(null);
            await Task.Delay(2000);
        }
    }

    public void AddBot(BotConfig config)
    {
        var tab = new BotTabViewModel(new BotEngine(config));
        Bots.Add(tab);
        SelectedBot = tab;
        SaveAll();
    }

    public async void RemoveBot(BotTabViewModel tab)
    {
        Bots.Remove(tab);
        if (SelectedBot == tab)
        {
            SelectedBot = Bots.Count > 0 ? Bots[0] : null;
        }

        await tab.DisposeAsync();
        SaveAll();
    }

    public void SaveAll()
    {
        _store.Save(Bots.Select(b => b.Config).ToList());
        RefreshAnyBotHasClanEnabled();
    }
}
