namespace Invigoration.App.Models;

/// <summary>
/// Old-school Battle.net-style presence, unified across classic BNCS
/// (derived from FriendStatus/FriendLocation flags) and Stimpak's own
/// Presence enum, so one icon set can represent a friend/roster entry from
/// either source.
/// </summary>
public enum PresenceState
{
    Offline,
    Available,
    Away,
    DoNotDisturb,
    InGame,
}
