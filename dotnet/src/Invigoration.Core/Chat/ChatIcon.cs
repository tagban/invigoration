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
    /// The product/game icon key. StarCraft: Brood War (PXES/"SEXP") is
    /// mapped to the plain StarCraft icon for now — SC:R self-identifies on
    /// the wire as plain Brood War with no distinct product code, and there's
    /// no dedicated icon to distinguish it, so it defaults to the standard
    /// StarCraft icon rather than a custom badge.
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
            "RATS" or "PXES" => "sc",
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
}
