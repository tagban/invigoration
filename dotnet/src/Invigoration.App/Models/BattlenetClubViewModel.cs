using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.Core.Sc2;

namespace Invigoration.App.Models;

/// <summary>One StarCraft II clan or group in the Clans list. Its members are folded away until opened with +.</summary>
public sealed partial class BattlenetClubViewModel(BattlenetClub club) : ObservableObject
{
    public BattlenetClub Club { get; } = club;

    public string Title => Club.Tag is { Length: > 0 } tag ? $"<{tag}> {Club.Name}" : Club.Name;

    /// <summary>The members, highest rank first; empty until Battle.net has sent them.</summary>
    public IReadOnlyList<BattlenetClubMember> Members => Club.Roster;

    public bool HasMembers => Club.Roster.Count > 0;

    public bool HasDescription => Club.Description.Length > 0;

    /// <summary>Whether the member list is showing (the + / − button).</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ExpandGlyph))]
    [NotifyPropertyChangedFor(nameof(ShowsMembers))]
    public partial bool IsExpanded { get; set; }

    public string ExpandGlyph => IsExpanded ? "−" : "+";

    public bool ShowsMembers => IsExpanded && HasMembers;

    public string Details =>
        $"{(Club.IsClan ? "Clan" : "Group")} · you: {Club.Rank} · {Club.Members} member{(Club.Members == 1 ? "" : "s")}" +
        (Club.Online > 0 ? $" · {Club.Online} online" : "");
}
