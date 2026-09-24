using Invigoration.Core.Chat;
using Stimpak;

namespace Invigoration.Core.Sc2;

/// <summary>
/// A chat event Stimpak's event types have no room for, such as SC:R's server notices and emotes,
/// passed through a native client's event stream so it's handled in order with everything else.
/// </summary>
public sealed record NativeChatEvent(ChatEvent Event) : SC2Event;

/// <summary>A native client's whole friends list, already in the Friends tab's terms. Replaces the previous one.</summary>
public sealed record NativeFriendsEvent(IReadOnlyList<FriendEntry> Friends) : SC2Event;

/// <summary>A Battle.net friend request. <see cref="Sent"/>: one this bot sent, rather than one to answer.</summary>
public sealed record FriendInvitation(ulong Id, string BattleTag, bool Sent);

/// <summary>A native client's pending friend requests, all of them. Replaces the previous list.</summary>
public sealed record NativeInvitationsEvent(IReadOnlyList<FriendInvitation> Invitations) : SC2Event;

/// <summary>A native client that can change the Battle.net friends list (SC:R so far; SC2's commands aren't mapped).</summary>
public interface IBattlenetFriendsClient
{
    /// <summary>Sends a friend request to a BattleTag.</summary>
    void AddFriend(string battleTag);

    /// <summary>Removes a Battle.net friend, by the BattleTag the Friends list shows.</summary>
    void RemoveFriend(string battleTag);

    void AnswerInvitation(ulong invitationId, bool accept);
}
