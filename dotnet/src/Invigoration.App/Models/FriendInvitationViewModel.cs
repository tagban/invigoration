namespace Invigoration.App.Models;

/// <summary>A pending Battle.net friend request on the Friends tab. <see cref="Sent"/>: one this bot sent, which only the other side can answer.</summary>
public sealed record FriendInvitationViewModel(ulong Id, string BattleTag, bool Sent)
{
    public bool CanAnswer => !Sent;

    public string Caption => Sent ? $"Request sent to {BattleTag}" : $"{BattleTag} wants to be friends";
}
