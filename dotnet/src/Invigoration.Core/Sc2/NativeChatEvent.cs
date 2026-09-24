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
