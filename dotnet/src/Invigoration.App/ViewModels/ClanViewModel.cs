using System.Collections.ObjectModel;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.Models;
using Invigoration.Core.Chat;
using Invigoration.Core.Clan;
using Invigoration.Core.Protocol;

namespace Invigoration.App.ViewModels;

/// <summary>
/// Manages the shared (cross-bot) "seen list" — everyone the bot has ever
/// observed talking, not just formally added clan members (see
/// ClanMember.IsClanMember). Edits apply directly to ClanRosterStore's
/// in-memory list as you type — Save persists them to disk, Close without
/// saving just leaves the in-memory list changed for this run (matches how
/// the roster already behaves for chat-command edits, which also mutate the
/// live list immediately).
/// </summary>
public partial class ClanViewModel : ObservableObject
{
    public ObservableCollection<ClanMemberViewModel> Members { get; }

    /// <summary>What the ItemsControl actually binds to — Members filtered by SearchText/ShowOnlyClanMembers, rebuilt whenever either changes.</summary>
    public ObservableCollection<ClanMemberViewModel> FilteredMembers { get; } = [];

    /// <summary>Just the rank names, for the per-member Rank dropdown — see Clan &gt; Manage Ranks to add/remove/rename ranks themselves.</summary>
    public ObservableCollection<string> AvailableRankNames { get; } = [];

    [ObservableProperty]
    public partial string SearchText { get; set; } = "";

    /// <summary>When true, the seen-list-wide "everyone ever observed" entries are hidden, leaving just formally added/promoted clan members.</summary>
    [ObservableProperty]
    public partial bool ShowOnlyClanMembers { get; set; }

    public ClanViewModel()
    {
        Members = new ObservableCollection<ClanMemberViewModel>(
            ClanRosterStore.Members.Select(m => new ClanMemberViewModel(m)));
        foreach (var rank in ClanRankStore.Ranks)
        {
            AvailableRankNames.Add(rank.Name);
        }

        RefreshFilter();
    }

    partial void OnSearchTextChanged(string value) => RefreshFilter();

    partial void OnShowOnlyClanMembersChanged(bool value) => RefreshFilter();

    private void RefreshFilter()
    {
        var query = SearchText.Trim();
        IEnumerable<ClanMemberViewModel> filtered = Members;

        if (ShowOnlyClanMembers)
        {
            filtered = filtered.Where(m => m.IsClanMember);
        }

        if (query.Length > 0)
        {
            filtered = filtered.Where(m =>
                m.Name.Contains(query, StringComparison.OrdinalIgnoreCase) ||
                m.NickName.Contains(query, StringComparison.OrdinalIgnoreCase) ||
                m.AliasesText.Contains(query, StringComparison.OrdinalIgnoreCase));
        }

        FilteredMembers.Clear();
        foreach (var member in filtered)
        {
            FilteredMembers.Add(member);
        }
    }

    /// <summary>Adds a new formal clan member (IsClanMember = true, distinct from someone auto-tracked just from chatting) and expands it immediately for editing.</summary>
    [RelayCommand]
    private void AddMember()
    {
        var member = new ClanMember { Name = "New Member", IsClanMember = true };
        ClanRosterStore.Members.Add(member);
        var vm = new ClanMemberViewModel(member) { IsExpanded = true };
        Members.Add(vm);
        RefreshFilter();
    }

    [RelayCommand]
    private void RemoveMember(ClanMemberViewModel member)
    {
        ClanRosterStore.Members.Remove(member.Member);
        Members.Remove(member);
        RefreshFilter();
    }

    [RelayCommand]
    private void Save() => ClanRosterStore.Save();
}

/// <summary>Wraps a ClanMember with a comma-separated Aliases editor and a collapsed/expanded row state for the seen-list UI.</summary>
public partial class ClanMemberViewModel : ObservableObject
{
    public ClanMember Member { get; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(DisplayName))]
    public partial string Name { get; set; }

    /// <summary>Freeform personal label (e.g. "John") — display/search only, never used to match a speaking user. Falls back to Name when blank.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(DisplayName))]
    public partial string NickName { get; set; }

    [ObservableProperty]
    public partial string Rank { get; set; }

    [ObservableProperty]
    public partial string AliasesText { get; set; }

    [ObservableProperty]
    public partial string Notes { get; set; }

    [ObservableProperty]
    public partial decimal TriviaScore { get; set; }

    [ObservableProperty]
    public partial bool IsClanMember { get; set; }

    /// <summary>Collapsed by default — a row shows just DisplayName/Rank/last-seen-game icon until expanded to the full editor.</summary>
    [ObservableProperty]
    public partial bool IsExpanded { get; set; }

    /// <summary>Free-form platform labels (e.g. "Classic, SC2, SC:R") — organizational only, doesn't affect matching.</summary>
    [ObservableProperty]
    public partial string PlatformsText { get; set; }

    /// <summary>What the collapsed row (and search) actually shows as the person's name — NickName if set, otherwise the Battle.net account Name.</summary>
    public string DisplayName => string.IsNullOrWhiteSpace(NickName) ? Name : NickName;

    /// <summary>Read-only — set by the bot as it observes chat, not hand-edited here.</summary>
    public string LastSeenText => Member.LastSeenUtc is { } seenUtc
        ? seenUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm")
        : "Never seen";

    /// <summary>Read-only — the Battle.net server this member was last observed on.</summary>
    public string LastSeenServerText => string.IsNullOrEmpty(Member.LastSeenServer) ? "Unknown server" : Member.LastSeenServer;

    /// <summary>Read-only — the display name of the game this member was last seen playing (e.g. "Diablo II: Lord of Destruction"), not just the bare wire code the icon is driven from.</summary>
    public string LastSeenGameText => string.IsNullOrEmpty(Member.LastSeenProduct)
        ? "Unknown game"
        : BncsProduct.GetDisplayName(Member.LastSeenProduct);

    /// <summary>Icon for the game this member was last seen playing, for the collapsed row — null if never observed with a known product.</summary>
    public Bitmap? LastSeenProductIconImage => string.IsNullOrEmpty(Member.LastSeenProduct)
        ? null
        : GameIconLoader.Get(ChatIcon.GetProductIconKey(Member.LastSeenProduct));

    /// <summary>Null (no tooltip at all, rather than an empty box) when there's nothing to show — bound to the collapsed row's own ToolTip.Tip so hovering a member with notes surfaces them without needing to expand the row.</summary>
    public string? NotesTooltip => string.IsNullOrWhiteSpace(Notes) ? null : Notes;

    public ClanMemberViewModel(ClanMember member)
    {
        Member = member;
        Name = member.Name;
        NickName = member.NickName;
        Rank = member.Rank;
        AliasesText = string.Join(", ", member.Aliases);
        Notes = member.Notes;
        TriviaScore = (decimal)member.TriviaScore;
        IsClanMember = member.IsClanMember;
        PlatformsText = string.Join(", ", member.Platforms);
    }

    partial void OnNameChanged(string value) => Member.Name = value;

    partial void OnNickNameChanged(string value) => Member.NickName = value;

    partial void OnRankChanged(string value) => Member.Rank = value;

    partial void OnAliasesTextChanged(string value) =>
        Member.Aliases = value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();

    partial void OnNotesChanged(string value)
    {
        Member.Notes = value;
        OnPropertyChanged(nameof(NotesTooltip));
    }

    partial void OnTriviaScoreChanged(decimal value) => Member.TriviaScore = (double)value;

    partial void OnIsClanMemberChanged(bool value) => Member.IsClanMember = value;

    partial void OnPlatformsTextChanged(string value) =>
        Member.Platforms = value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();

    [RelayCommand]
    private void ToggleExpanded() => IsExpanded = !IsExpanded;
}
