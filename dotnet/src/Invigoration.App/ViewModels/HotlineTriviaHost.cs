using Invigoration.Core.Tracking;
using Invigoration.Core.Trivia;

namespace Invigoration.App.ViewModels;

/// <summary>
/// Hotline's ITriviaHost implementation — lets one HotlineSessionViewModel run the exact same
/// Trivia.TriviaEngine round-runner Battle.net bots use (see BotEngine.Trivia.cs's own
/// implementation), instead of a second, drifting copy. Scores/last-seen go to
/// Tracking.ProtocolUserTrackingStore — deliberately NOT Clan.ClanRosterStore/the Battle.net "CLAN"
/// list, per explicit request; these are Hotline users, never clan members. A round here never
/// leaves the single server it started on: there's no Hotline equivalent of BotConfig.TriviaGroup,
/// so BroadcastAsync/AnnounceStartedAsync always just send to this one session's own chat.
/// </summary>
public sealed class HotlineTriviaHost(HotlineSessionViewModel session) : ITriviaHost
{
    private const string Protocol = "Hotline";

    private string ServerTag => $"{session.Host}:{session.Port}";

    public bool TriviaFeatureEnabled => session.TriviaEnabled;

    // No per-server override for these exists yet — matches BotConfig's own defaults
    // (TriviaPointsBeforeFirstHint/AfterFirstHint/AfterSecondHint).
    public double PointsBeforeFirstHint => 1.25;

    public double PointsAfterFirstHint => 1.0;

    public double PointsAfterSecondHint => 0.75;

    public Task BroadcastAsync(string text) => session.SendChatWithEffectsAsync(text);

    /// <summary>No SC2-style sub-channel targeting to freeze here — a Hotline session has exactly one chat stream.</summary>
    public void OnRoundStarting()
    {
    }

    /// <summary>There's only ever one place a Hotline round's messages can go, so the "started" confirmation is no different from any other round message.</summary>
    public Task AnnounceStartedAsync(string text, Func<string, Task> reply) => BroadcastAsync(text);

    public string RecordScore(string username, double points)
    {
        var newScore = ProtocolUserTrackingStore.AddScore(username, Protocol, ServerTag, points);
        return $" (+{points.ToString("0.##")}) Your score is {newScore.ToString("0.##")}.";
    }

    public string FormatLeaderboard()
    {
        var leaders = ProtocolUserTrackingStore.GetLeaderboard(Protocol, ServerTag);
        return leaders.Count == 0 ? "" : string.Join(" ", leaders.Select(u => $"{u.Name}({u.TriviaScore})"));
    }

    public void LogParseErrors(IReadOnlyList<string> errors) =>
        session.AppendDebugMessage($"Trivia: skipped {errors.Count} unparsable line(s): {string.Join("; ", errors.Take(5))}");
}
