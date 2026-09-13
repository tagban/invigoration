namespace Invigoration.Core.Chat;

/// <summary>Which named palette a bot's chat log renders with, matching bnubot's ColorScheme options.</summary>
public enum ChatColorScheme
{
    /// <summary>Invigoration's own signature palette — this port's original default.</summary>
    Invigoration,

    StarCraft,

    DiabloII,

    /// <summary>Every role hand-picked by the user, stored in BotConfig.CustomColors.</summary>
    Custom,

    /// <summary>Warm parchment and gold on dark oak, for the Warcraft theme. Invigoration's own, not ported from bnubot. Added after Custom so existing saved values keep their meaning.</summary>
    Warcraft,
}
