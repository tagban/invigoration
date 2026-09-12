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
    ///
    /// <paramref name="flags"/> (default 0, i.e. no status icon) lets Diablo pick one of four
    /// "how far have they gotten" badges (classic.battle.net/info/icons.shtml's "Diablo Icons"
    /// section: 0/1/2/3 red dots for no/Normal/Nightmare/Hell clears) instead of the flat generic
    /// logo — but only when the user has no status badge of their own (GetStatusIconKey(flags) ==
    /// ""), matching the explicit request to show "game info" in that slot specifically when
    /// there's nothing more important (a rank/mute/guest badge) to show there instead. Every
    /// other call site either doesn't have live flags to pass (clan/friend rows) or doesn't care,
    /// and the default leaves them with the same "always show the rank badge if one applies"
    /// behavior — there's no flags-driven badge to lose priority to there anyway.
    /// </summary>
    public static string GetProductIconKey(string statString, uint flags = 0)
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
        var noStatusIcon = GetStatusIconKey(flags) == "";

        if (product == "LTRD" && noStatusIcon && TryGetDiabloDots(statString) is { } dots)
        {
            return $"diablo-dot{dots}";
        }

        // StarCraft/Brood War/SC-Japanese share one generic win-based rank badge on real
        // Battle.net (classic.battle.net/info/icons.shtml lists it once, not per-product) —
        // pixel-exact sprites extracted from the real icons_STAR.bni (bnetdocs.org/document/25/
        // icons-bni), not fabricated. "sc-stars0" is deliberately reused for both 0 wins AND a
        // ranked player (confirmed): once a player has an actual ladder rating, the real client
        // drops the win-count star display entirely in favor of stamping the numeric rating on
        // this same plain plate — see GetLadderScore, which the UI overlays as text.
        if (product is "RATS" or "PXES" or "RTSJ" && noStatusIcon && TryGetIconClassStats(statString, out var scRating, out var scWins))
        {
            return scRating > 0 ? "sc-stars0" : $"sc-stars{Math.Clamp(scWins, 0, 10)}";
        }

        // Warcraft II BNE: half an axe per win up to 4 axes (8 wins), a single sword at 9 wins,
        // two swords from 10 wins until the player is actually ladder-rated — at which point,
        // same as StarCraft above, the badge switches to a plain plate with the rating stamped on
        // it instead of axes/swords. Real icons_STAR.bni-family sprites extracted from
        // icons_W2BN.bni via the same bnetdocs.org archive.
        if (product == "NB2W" && noStatusIcon && TryGetIconClassStats(statString, out var w2Rating, out var w2Wins))
        {
            if (w2Rating > 0)
            {
                return "war2-ranked";
            }

            return w2Wins switch
            {
                <= 8 => $"war2-axes{w2Wins}",
                9 => "war2-sword1",
                _ => "war2-sword2",
            };
        }

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
            "TAHC" => "chat",
            _ => "",
        };
    }

    /// <summary>
    /// The ladder rating ("score") to stamp as text over a StarCraft-family/Warcraft II win badge
    /// once a player has one — a value of 0 means "not yet ladder-rated" (matches
    /// ParseIconClassStats' own check), in which case there's nothing to stamp and this returns
    /// null. Not applicable (also null) to Diablo, whose rank badge (GetProductIconKey's dot
    /// picker) has no equivalent numeric score to show.
    /// </summary>
    public static int? GetLadderScore(string statString)
    {
        var product = statString.Length >= 4 ? statString[..4] : statString;
        if (product is not ("RATS" or "PXES" or "RTSJ" or "NB2W"))
        {
            return null;
        }

        return TryGetIconClassStats(statString, out var rating, out _) && rating > 0 ? rating : null;
    }

    /// <summary>
    /// The number of "red dots" (0-3) a Diablo statstring's icon-class stats report — same
    /// space-split wire layout StatStringParser.ParseDiabloClassicStats reads, duplicated here
    /// rather than shared since that method builds a sentence and has no reason to expose a
    /// structured field. Field order (bnetdocs.org/document/18/chat-statstrings, confirmed
    /// 2026-09-11): Level, Class, Dots, Strength, Magic, Dexterity, Vitality, Gold, Spawned — dots
    /// is index 2, not index 1 (that was a real bug here and in ParseDiabloClassicStats, both
    /// fixed together). Null if the statstring doesn't parse as Diablo icon-class stats at all
    /// (e.g. an Open Character with no stats yet) — GetProductIconKey falls back to the flat
    /// "diablo" icon in that case.
    /// </summary>
    private static int? TryGetDiabloDots(string statString)
    {
        var values = statString.Length > 5 ? statString[5..].Split(' ') : [];
        return values.Length == 9 && int.TryParse(values[2], out var dots) ? Math.Clamp(dots, 0, 3) : null;
    }

    /// <summary>
    /// Rating (index 0) and wins (index 2) out of the same "icon class stats" wire layout
    /// StatStringParser.ParseIconClassStats reads for StarCraft/Brood War/SC-Japanese/Warcraft II
    /// — duplicated here for the same reason as TryGetDiabloDots (that method formats a sentence
    /// and has no reason to expose structured fields). False (both out params 0) if the
    /// statstring doesn't parse as icon-class stats at all.
    /// </summary>
    private static bool TryGetIconClassStats(string statString, out int rating, out int wins)
    {
        var values = statString.Length > 5 ? statString[5..].Split(' ') : [];
        if (values.Length == 9 && int.TryParse(values[0], out rating) && int.TryParse(values[2], out wins))
        {
            return true;
        }

        rating = 0;
        wins = 0;
        return false;
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

        // "Special guest" (BLIZZARD_GUEST) — the sunglasses badge, kept distinct from Speaker
        // (mega.png) per explicit request, since the two had previously been folded together
        // under one "Speaker / VIP" label even though they're different flags/icons on real
        // Battle.net. See ChatPalette, which already treats this same flag as "Guest" for text
        // color — this is the matching icon-badge half of that, previously unhandled here.
        // guest.png is real Blizzard art (index 4 of the shared 0-7 status-icon prefix in both
        // icons_STAR.bni and icons_W2BN.bni — bnetdocs.org/document/25/icons-bni), replacing an
        // earlier hand-drawn placeholder once the real file was found.
        if (uflags.HasFlag(UserFlags.Special))
        {
            return "guest";
        }

        if (uflags.HasFlag(UserFlags.Squelched))
        {
            return "ignore";
        }

        return "";
    }

    /// <summary>
    /// True for any rank badge that should sort to the top of a channel's user list — Blizzard
    /// rep, Admin, Operator ("has a gavel"), or Speaker — matching classic Battle.net's own
    /// "moderators, then everyone else" ordering per user request. Squelched and Special guest
    /// deliberately aren't included: one's a punishment marker, the other's just a badge, and
    /// neither is a rank that should float someone to the top.
    /// </summary>
    public static bool IsPrivileged(uint flags)
    {
        var uflags = (UserFlags)flags;
        return uflags.HasFlag(UserFlags.Blizzard) || uflags.HasFlag(UserFlags.Admin) ||
               uflags.HasFlag(UserFlags.Operator) || uflags.HasFlag(UserFlags.Speaker);
    }
}
