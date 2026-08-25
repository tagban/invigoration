namespace Invigoration.Core.Music;

/// <summary>
/// One shared YouTube Music player for the whole app, not per-bot — every bot's `!skip`/
/// `!thumbsup`/`!thumbsdown`/`!nowplaying` command controls the same embedded window regardless
/// of which bot tab it came from. Simpler than Trivia.TriviaGroupRegistry's per-group-session
/// pattern (see BotEngine.cs's TriviaGroupRegistry usage): there's only ever one player, and a
/// command's result only needs to reply on the channel it came from, not broadcast to every bot.
/// Set by YouTubeMusicWindow when it opens (App layer — Core stays UI-agnostic, same bridge
/// pattern as BotEngine.Sc2ChallengeHandler) and cleared back to null when it closes, so a
/// command sent while the window isn't open gets a clear "not open" reply instead of a
/// NullReferenceException.
/// </summary>
public static class MusicPlayerRegistry
{
    public static IMusicPlayerController? Controller { get; set; }
}
