using System.Collections.Concurrent;
using Invigoration.Core.Chat;

namespace Invigoration.Core.Sc2;

/// <summary>
/// Which game each StarCraft: Remastered chat member is on, by name. SC:R chat is classic chat, and
/// each member carries a classic product code ("program_id": SEXP, DRTL, ...), but Stimpak's Person
/// has nowhere to put it, so the user list looks it up here. Keyed by name across every bot: a name
/// is on one game at a time.
/// </summary>
public static class NativeMemberProducts
{
    private static readonly ConcurrentDictionary<string, string> ByName = new(StringComparer.OrdinalIgnoreCase);
    private static readonly ConcurrentDictionary<string, string> BattleTags = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>An SC:R member's BattleTag, which Battle.net sends with each channel member.</summary>
    public static void SetBattleTag(string name, string battleTag)
    {
        if (name.Length > 0 && battleTag.Contains('#'))
        {
            BattleTags[name] = battleTag;
        }
    }

    public static string? BattleTagFor(string name) => BattleTags.TryGetValue(name, out var tag) ? tag : null;

    public static void Set(string name, string programId)
    {
        if (name.Length > 0 && programId.Length == 4)
        {
            ByName[name] = programId;
        }
    }

    /// <summary>Whether this name is a StarCraft: Remastered chat member, which has no presence to show.</summary>
    public static bool IsKnown(string name) => ByName.ContainsKey(name);

    /// <summary>The icon key for a member's game, or null when none is known (every SC2 user).</summary>
    public static string? IconKeyFor(string name) =>
        ByName.TryGetValue(name, out var program)
            // ChatIcon takes the wire-order (reversed) code a classic statstring starts with.
            ? ChatIcon.GetProductIconKey(new string(program.Reverse().ToArray()), showLadder: false) is { Length: > 0 } key ? key : null
            : null;
}
