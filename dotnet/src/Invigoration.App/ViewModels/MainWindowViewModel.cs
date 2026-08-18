using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.Core;
using Invigoration.Core.Config;

namespace Invigoration.App.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    private readonly ConfigStore _store = new();

    public ObservableCollection<BotTabViewModel> Bots { get; } = [];

    [ObservableProperty]
    public partial BotTabViewModel? SelectedBot { get; set; }

    public MainWindowViewModel()
    {
        foreach (var config in _store.Load())
        {
            Bots.Add(new BotTabViewModel(new BotEngine(config)));
        }

        SelectedBot = Bots.Count > 0 ? Bots[0] : null;
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

    public void SaveAll() => _store.Save(Bots.Select(b => b.Config).ToList());
}
