namespace Invigoration.Core.Trivia;

/// <summary>
/// What TriviaEngine needs from whichever chat client is hosting a round — implemented by both
/// BotEngine (Battle.net, via BotEngine.Trivia.cs) and Hotline's own HotlineTriviaHost, so the
/// exact same round-running/hint-timing/scoring logic (TriviaEngine) works identically on both
/// instead of two independently-maintained copies. Deliberately thin: everything protocol-specific
/// (how a message actually gets sent, where scores live) is behind this seam; TriviaEngine itself
/// never touches BNCS or Hotline directly.
/// </summary>
public interface ITriviaHost
{
    bool TriviaFeatureEnabled { get; }

    double PointsBeforeFirstHint { get; }

    double PointsAfterFirstHint { get; }

    double PointsAfterSecondHint { get; }

    /// <summary>Sends one trivia round message to everywhere this round is actually visible — just this host's own chat for a Hotline session (a round never leaves the single server it was started on), or this bot's channel plus every TriviaGroup peer for BotEngine.</summary>
    Task BroadcastAsync(string text);

    /// <summary>Called once, right after a round's question pool is set (TriviaSession.Start) and before the first message goes out — lets BotEngine freeze its SC2 sub-channel target for the round the same way it always has (_sc2TriviaChannelIndex); a no-op for Hotline, which has no sub-channels.</summary>
    void OnRoundStarting();

    /// <summary>
    /// Announces the "Trivia started..." confirmation — the one round message whose target can
    /// differ from BroadcastAsync's usual one: BotEngine replies to wherever the "!trivia on" was
    /// typed (possibly a whisper) when no TriviaGroup peers need to see it, but broadcasts it like
    /// any other round message when they do. Hotline always just broadcasts (there's only ever one
    /// place to send).
    /// </summary>
    Task AnnounceStartedAsync(string text, Func<string, Task> reply);

    /// <summary>
    /// Records a correct answer's score against whatever this host's own user-tracking store is
    /// (Clan.ClanRosterStore for BotEngine, Tracking.ProtocolUserTrackingStore for Hotline) and
    /// returns the short " (+N) Your score is X." fragment (leading space, no leading "!" — the
    /// caller already appends this straight after "!") to append to the winner announcement, or ""
    /// if this host doesn't track a score for the given username (matches BotEngine's existing
    /// behavior: an unrecognized/never-added Battle.net name plays but never scores).
    /// </summary>
    string RecordScore(string username, double points);

    /// <summary>The "!trivia score" / "/trivia score" leaderboard body — "Name(Score) Name(Score) ..." for up to the top 10 scorers, or "" if nobody has scored yet.</summary>
    string FormatLeaderboard();

    /// <summary>A malformed line (or lines) in the trivia question packs — a local diagnostic only, never sent to chat.</summary>
    void LogParseErrors(IReadOnlyList<string> errors);
}
