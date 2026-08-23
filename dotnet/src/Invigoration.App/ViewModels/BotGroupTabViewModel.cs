using System.Collections.ObjectModel;
using Avalonia.Media;
using CommunityToolkit.Mvvm.ComponentModel;

namespace Invigoration.App.ViewModels;

/// <summary>
/// A named group of bots (BotConfig.TabGroup) collapsed into one top-level tab — declutters the
/// tab strip when several bots share a group (e.g. all connected to the same server). Renders
/// in MainWindowViewModel.TopLevelTabs' shared header ItemTemplate the same duck-typed way
/// GlobalWhispersTabViewModel does (Title/HighlightBrush by name, not a shared base type), and
/// its own content is a plain nested TabControl over Bots — the same pattern BotTabView.axaml
/// already uses for a Stimpak-backed bot's own per-channel sub-tabs.
/// </summary>
public sealed partial class BotGroupTabViewModel : ViewModelBase
{
    public string Title { get; }

    /// <summary>Borrows the first member's scheme color rather than picking one of its own — a group has no palette of its own to draw from.</summary>
    public IBrush HighlightBrush { get; }

    /// <summary>Same normal-tab header look as a plain BotTabViewModel — see GlobalWhispersTabViewModel's matching properties for why the Whispers pseudo-tab overrides both.</summary>
    public double HeaderFontSize => 13;

    public IBrush HeaderForeground => Brushes.White;

    public ObservableCollection<BotTabViewModel> Bots { get; }

    [ObservableProperty]
    public partial BotTabViewModel? SelectedBot { get; set; }

    /// <summary>True while any member bot not currently in view has something unread — see the constructor, which subscribes to each member's HasUnread and re-raises this whenever one changes.</summary>
    public bool HasUnread => Bots.Any(b => b.HasUnread);

    public BotGroupTabViewModel(string groupName, IEnumerable<BotTabViewModel> members)
    {
        Title = groupName;
        Bots = new ObservableCollection<BotTabViewModel>(members);
        HighlightBrush = Bots.Count > 0 ? Bots[0].HighlightBrush : Brushes.Gray;
        SelectedBot = Bots.Count > 0 ? Bots[0] : null;

        foreach (var bot in Bots)
        {
            bot.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(BotTabViewModel.HasUnread))
                {
                    OnPropertyChanged(nameof(HasUnread));
                }
            };
        }
    }
}
