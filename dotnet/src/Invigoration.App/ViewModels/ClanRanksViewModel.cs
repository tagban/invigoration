using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.Core.Clan;
using Invigoration.Core.Commands;

namespace Invigoration.App.ViewModels;

/// <summary>
/// Manages the shared (cross-bot) list of predefined ranks — each optionally
/// carrying auto-whisper/auto-kick/auto-ban behaviors and a set of commands
/// it grants. This replaces the old per-bot BotConfig.PermissionLevels
/// system: command access now lives directly on the rank, shared across
/// every bot instead of configured separately per bot.
/// </summary>
public partial class ClanRanksViewModel : ObservableObject
{
    public ObservableCollection<ClanRankViewModel> Ranks { get; }

    public ClanRanksViewModel()
    {
        Ranks = new ObservableCollection<ClanRankViewModel>(ClanRankStore.Ranks.Select(r => new ClanRankViewModel(r)));
    }

    [RelayCommand]
    private void AddRank()
    {
        var rank = new ClanRank { Name = "New Rank" };
        ClanRankStore.Ranks.Add(rank);
        Ranks.Add(new ClanRankViewModel(rank));
    }

    [RelayCommand]
    private void RemoveRank(ClanRankViewModel rank)
    {
        ClanRankStore.Ranks.Remove(rank.Rank);
        Ranks.Remove(rank);
    }

    [RelayCommand]
    private void Save() => ClanRankStore.Save();
}

/// <summary>Wraps a ClanRank for the "Manage Ranks" editor.</summary>
public partial class ClanRankViewModel : ObservableObject
{
    public ClanRank Rank { get; }

    [ObservableProperty]
    public partial string Name { get; set; }

    [ObservableProperty]
    public partial string AutoWhisperMessage { get; set; }

    [ObservableProperty]
    public partial AutoWhisperFrequency AutoWhisperFrequency { get; set; }

    [ObservableProperty]
    public partial bool AutoKick { get; set; }

    /// <summary>Optional reason sent with the auto-kick (e.g. "/kick username reason") — blank sends no reason.</summary>
    [ObservableProperty]
    public partial string AutoKickMessage { get; set; }

    [ObservableProperty]
    public partial bool AutoBan { get; set; }

    /// <summary>Optional reason sent with the auto-ban (e.g. "/ban username reason") — blank sends no reason.</summary>
    [ObservableProperty]
    public partial string AutoBanMessage { get; set; }

    public IReadOnlyList<AutoWhisperFrequency> AvailableFrequencies { get; } = Enum.GetValues<AutoWhisperFrequency>();

    /// <summary>Per-command checklist granting this rank access — mirrors the old PermissionLevel command checklist.</summary>
    public ObservableCollection<RankCommandGrantViewModel> Commands { get; }

    /// <summary>Collapsed by default — a row shows just Name/SummaryText until expanded to the full editor, matching the Seen List's member rows.</summary>
    [ObservableProperty]
    public partial bool IsExpanded { get; set; }

    /// <summary>One-line dim subtitle for the collapsed row — which behaviors are actually turned on, so you don't have to expand every rank just to see what it does.</summary>
    [ObservableProperty]
    public partial string SummaryText { get; set; } = "";

    public ClanRankViewModel(ClanRank rank)
    {
        Rank = rank;
        Name = rank.Name;
        AutoWhisperMessage = rank.AutoWhisperMessage;
        AutoWhisperFrequency = rank.AutoWhisperFrequency;
        AutoKick = rank.AutoKick;
        AutoKickMessage = rank.AutoKickMessage;
        AutoBan = rank.AutoBan;
        AutoBanMessage = rank.AutoBanMessage;
        Commands = new ObservableCollection<RankCommandGrantViewModel>(
            CommandCatalog.Entries.Select(e => new RankCommandGrantViewModel(e, rank, RefreshSummaryText)));
        RefreshSummaryText();
    }

    partial void OnNameChanged(string value) => Rank.Name = value;

    partial void OnAutoWhisperMessageChanged(string value)
    {
        Rank.AutoWhisperMessage = value;
        RefreshSummaryText();
    }

    partial void OnAutoWhisperFrequencyChanged(AutoWhisperFrequency value)
    {
        Rank.AutoWhisperFrequency = value;
        RefreshSummaryText();
    }

    partial void OnAutoKickChanged(bool value)
    {
        Rank.AutoKick = value;
        RefreshSummaryText();
    }

    partial void OnAutoKickMessageChanged(string value) => Rank.AutoKickMessage = value;

    partial void OnAutoBanChanged(bool value)
    {
        Rank.AutoBan = value;
        RefreshSummaryText();
    }

    partial void OnAutoBanMessageChanged(string value) => Rank.AutoBanMessage = value;

    private void RefreshSummaryText()
    {
        var parts = new List<string>();
        var grantedCount = Commands?.Count(c => c.IsGranted) ?? 0;
        if (grantedCount > 0)
        {
            parts.Add($"{grantedCount} command{(grantedCount == 1 ? "" : "s")}");
        }

        if (!string.IsNullOrWhiteSpace(AutoWhisperMessage))
        {
            parts.Add($"Auto-whisper ({AutoWhisperFrequency})");
        }

        if (AutoBan)
        {
            parts.Add("Auto-ban");
        }
        else if (AutoKick)
        {
            parts.Add("Auto-kick");
        }

        SummaryText = parts.Count > 0 ? string.Join(" · ", parts) : "No commands or behaviors set";
    }

    [RelayCommand]
    private void ToggleExpanded() => IsExpanded = !IsExpanded;
}

/// <summary>One checkbox in a rank's command checklist.</summary>
public partial class RankCommandGrantViewModel : ObservableObject
{
    private readonly ClanRank _rank;
    private readonly string _canonicalName;
    private readonly Action _onGrantedChanged;

    public string DisplayName { get; }

    [ObservableProperty]
    public partial bool IsGranted { get; set; }

    public RankCommandGrantViewModel(CommandCatalogEntry entry, ClanRank rank, Action onGrantedChanged)
    {
        _rank = rank;
        _canonicalName = entry.CanonicalName;
        _onGrantedChanged = onGrantedChanged;
        DisplayName = entry.DisplayName;
        IsGranted = rank.AllowedCommands.Contains(entry.CanonicalName);
    }

    partial void OnIsGrantedChanged(bool value)
    {
        if (value)
        {
            if (!_rank.AllowedCommands.Contains(_canonicalName))
            {
                _rank.AllowedCommands.Add(_canonicalName);
            }
        }
        else
        {
            _rank.AllowedCommands.Remove(_canonicalName);
        }

        _onGrantedChanged();
    }
}
