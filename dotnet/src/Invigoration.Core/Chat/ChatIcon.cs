namespace Invigoration.Core.Chat;

/// <summary>
/// Maps a user's flags/statstring to a Battle.net chat icon key, returned as
/// a filename (without extension) the UI resolves to an image. Mostly the
/// original classic.battle.net "chat icons" set (classic.battle.net/info/icons.shtml),
/// except the moderator/channel-operator badge uses a custom flat green gavel
/// with a transparent background (mod-gavel.png) instead — it reads cleanly
/// against the app's dark theme, unlike the original icon's opaque white tile.
/// Split into a product icon and an optional status/rank badge so a UI can
/// show both at once (e.g. product icon left, moderator badge right) instead
/// of the original bnetbot.cls behavior of picking one icon exclusively.
/// </summary>
public static class ChatIcon
{
    /// <summary>
    /// The product/game icon key. StarCraft: Brood War (PXES/"SEXP", wire-order reversed) now
    /// gets its own "scbw" badge — fixed 2026-08-24, was previously folded into plain "sc" (that
    /// was wrong: bnetdocs confirms PXES is Brood War's real product code, not a StarCraft one;
    /// SC:R self-identifying on the wire as plain Brood War, with no distinct product code of its
    /// own, correctly picks up the same Brood War badge here too — not a bug, since it really is
    /// Brood War-based).
    /// </summary>
    public static string GetProductIconKey(string statString)
    {
        // Not a wire-order product code at all: BotEngine.Sc2.cs stamps this literal
        // sentinel on every Stimpak-backed (SC2/SC:R/WC3:R) friend, since Stimpak's own
        // Friend/Person records don't carry a per-contact product code the way classic
        // BNCS's statstring does — every Stimpak contact gets the same icon today.
        if (statString == "sc2")
        {
            return "sc2";
        }

        var product = statString.Length >= 4 ? statString[..4] : statString;
        return product switch
        {
            "3RAW" or "PX3W" => "war3",
            "PX2D" => "d2exp",
            "VD2D" => "diablo2",
            "LTRD" => "diablo",
            "RHSD" => "dshr",
            "RATS" => "sc",
            "PXES" => "scbw",
            "RHSS" => "sware",
            "RTSJ" => "jsc",
            "NB2W" => "war2",
            _ => "",
        };
    }

    /// <summary>The status/rank badge icon key, or "" if the user has none of these flags.</summary>
    public static string GetStatusIconKey(uint flags)
    {
        var uflags = (UserFlags)flags;

        if (uflags.HasFlag(UserFlags.Blizzard))
        {
            return "blizz";
        }

        if (uflags.HasFlag(UserFlags.Admin))
        {
            return "sysop";
        }

        if (uflags.HasFlag(UserFlags.Operator))
        {
            return "mod-gavel";
        }

        if (uflags.HasFlag(UserFlags.Speaker))
        {
            return "mega";
        }

        if (uflags.HasFlag(UserFlags.Squelched))
        {
            return "ignore";
        }

        return "";
    }

    /// <summary>
    /// True for any rank badge that should sort to the top of a channel's user list — Blizzard
    /// rep, Admin, Operator ("has a gavel"), or Speaker/VIP — matching classic Battle.net's own
    /// "moderators, then everyone else" ordering per user request. Squelched deliberately isn't
    /// included: it's a punishment marker, not a rank, and shouldn't float someone to the top.
    /// </summary>
    public static bool IsPrivileged(uint flags)
    {
        var uflags = (UserFlags)flags;
        return uflags.HasFlag(UserFlags.Blizzard) || uflags.HasFlag(UserFlags.Admin) ||
               uflags.HasFlag(UserFlags.Operator) || uflags.HasFlag(UserFlags.Speaker);
    }
}
