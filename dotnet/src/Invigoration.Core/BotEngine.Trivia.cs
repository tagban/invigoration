using Invigoration.Core.Clan;
using Invigoration.Core.Trivia;

namespace Invigoration.Core;

/// <summary>
/// Chat trivia game, ported from BNU`Bot's TriviaEventHandler
/// (github.com/tagban/bnubot/tree/master/BNUBot/src/net/bnubot/bot/trivia) —
/// same on/off toggle, hint-reveal timing, and auto-shutoff after 10
/// consecutive unanswered questions in a row. The actual round-running logic
/// now lives in the shared <see cref="Trivia.TriviaEngine"/> (see its
/// remarks) so it's identical for a Hotline session's own trivia — this file
/// is just BotEngine's <see cref="ITriviaHost"/> implementation: how a round
/// message actually gets sent (BNCS/Stimpak), and where scores/points come
/// from (Config, Clan.ClanRosterStore).
///
/// Departs from the original in two places:
/// 1. BNU`Bot scores anyone who speaks automatically, because it keeps a
///    persistent account for every user it's ever seen. This port stores
///    trivia score on the existing clan-roster member record instead of a
///    separate always-on player database, so an unrecognized name plays but
///    never scores (see RecordScore).
/// 2. BNU`Bot's trivia is always per-connection. Here, when
///    <see cref="Config"/>.TriviaGroup is set, this bot shares one
///    <see cref="TriviaSession"/> (via <see cref="TriviaGroupRegistry"/>)
///    with every other bot naming the same group — e.g. a Warcraft II bot
///    and a StarCraft II bot run by the same person can host one combined
///    round, each relaying the question to its own channel and accepting
///    answers from its own players into the shared score/join state.
/// </summary>
public sealed partial class BotEngine : ITriviaHost
{
    private TriviaSession? _ownTrivia;
    private TriviaEngine? _triviaEngine;

    /// <summary>The active session: shared with this bot's TriviaGroup peers if one is set, otherwise a private one just for this bot. Resolved fresh each access since Config can be replaced after an edit.</summary>
    private TriviaSession _trivia => string.IsNullOrWhiteSpace(Config.TriviaGroup)
        ? (_ownTrivia ??= new TriviaSession())
        : TriviaGroupRegistry.GetSession(Config.TriviaGroup);

    private TriviaEngine TriviaEngineInstance => _triviaEngine ??= new TriviaEngine(this, () => _trivia);

    private Task HandleTriviaCommandAsync(string rest, Func<string, Task> reply) =>
        TriviaEngineInstance.HandleCommandAsync(rest, reply);

    bool ITriviaHost.TriviaFeatureEnabled => Config.TriviaFeatureEnabled;

    double ITriviaHost.PointsBeforeFirstHint => Config.TriviaPointsBeforeFirstHint;

    double ITriviaHost.PointsAfterFirstHint => Config.TriviaPointsAfterFirstHint;

    double ITriviaHost.PointsAfterSecondHint => Config.TriviaPointsAfterSecondHint;

    void ITriviaHost.OnRoundStarting() =>
        _sc2TriviaChannelIndex = Protocol.BncsProduct.IsStimpakBacked(Config.Product) ? _sc2ActiveChannelIndex : null;

    /// <summary>Matches the original: a non-grouped round's "started" confirmation goes only to wherever the command was typed (possibly a whisper), but a grouped round broadcasts it like any other round message so every peer's channel sees it too.</summary>
    Task ITriviaHost.AnnounceStartedAsync(string text, Func<string, Task> reply) =>
        string.IsNullOrWhiteSpace(Config.TriviaGroup) ? reply(text) : BroadcastTriviaMessageAsync(text);

    string ITriviaHost.RecordScore(string username, double points)
    {
        var member = ClanRosterStore.Find(username);
        if (member is null)
        {
            return "";
        }

        member.TriviaScore += points;
        ClanRosterStore.Save();
        return $" (+{points.ToString("0.##")}) Your score is {member.TriviaScore.ToString("0.##")}.";
    }

    string ITriviaHost.FormatLeaderboard()
    {
        var leaders = ClanRosterStore.Members
            .Where(m => m.TriviaScore != 0)
            .OrderByDescending(m => m.TriviaScore)
            .Take(10)
            .ToList();

        return leaders.Count == 0 ? "" : string.Join(" ", leaders.Select(m => $"{m.Name}({m.TriviaScore})"));
    }

    void ITriviaHost.LogParseErrors(IReadOnlyList<string> errors) =>
        LogInfo($"Trivia: skipped {errors.Count} unparsable line(s): {string.Join("; ", errors.Take(5))}");

    /// <summary>Sends a trivia round message to this bot's own channel, and to every other bot sharing its TriviaGroup, so linked channels all see the game. A dead/disconnected peer is logged and skipped rather than failing the whole broadcast.</summary>
    async Task ITriviaHost.BroadcastAsync(string text) => await BroadcastTriviaMessageAsync(text).ConfigureAwait(false);

    private async Task BroadcastTriviaMessageAsync(string text)
    {
        // Explicit channel, not the ambient active one: an operator switching sub-tabs
        // mid-round must not redirect the rest of this round's messages to a different
        // channel than the one it actually started in. No-op override (null) for BNCS.
        await SendChatCommandAsync(text, _sc2TriviaChannelIndex).ConfigureAwait(false);

        foreach (var peer in TriviaGroupRegistry.GetGroupPeers(Config.TriviaGroup, this))
        {
            try
            {
                await peer.SendChatCommandAsync(text).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                LogInfo($"Trivia: failed to relay a message to linked bot \"{peer.Config.DisplayName}\": {ex.Message}");
            }
        }
    }
}
