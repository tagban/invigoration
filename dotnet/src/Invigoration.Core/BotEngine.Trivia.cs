using Invigoration.Core.Clan;
using Invigoration.Core.Trivia;

namespace Invigoration.Core;

/// <summary>
/// Chat trivia game, ported from BNU`Bot's TriviaEventHandler
/// (github.com/tagban/bnubot/tree/master/BNUBot/src/net/bnubot/bot/trivia) —
/// same on/off toggle, hint-reveal timing (initial hint with the question,
/// more at +10s/+20s, 30s total per question), and auto-shutoff after 10
/// consecutive unanswered questions in a row.
///
/// Departs from the original in two places:
/// 1. BNU`Bot scores anyone who speaks automatically, because it keeps a
///    persistent account for every user it's ever seen. This port stores
///    trivia score on the existing clan-roster member record instead of a
///    separate always-on player database (per how
///    <see cref="ClanMember.TriviaScore"/> was scoped), so a first-time
///    player needs an explicit "!trivia join" — which both opts them into
///    having their chat scanned for answers and creates their roster entry.
/// 2. BNU`Bot's trivia is always per-connection. Here, when
///    <see cref="Config"/>.TriviaGroup is set, this bot shares one
///    <see cref="TriviaSession"/> (via <see cref="TriviaGroupRegistry"/>)
///    with every other bot naming the same group — e.g. a Warcraft II bot
///    and a StarCraft II bot run by the same person can host one combined
///    round, each relaying the question to its own channel and accepting
///    answers from its own players into the shared score/join state.
/// </summary>
public sealed partial class BotEngine
{
    private TriviaSession? _ownTrivia;
    private CancellationTokenSource? _triviaRoundCts;

    /// <summary>The active session: shared with this bot's TriviaGroup peers if one is set, otherwise a private one just for this bot. Resolved fresh each access since Config can be replaced after an edit.</summary>
    private TriviaSession _trivia => string.IsNullOrWhiteSpace(Config.TriviaGroup)
        ? (_ownTrivia ??= new TriviaSession())
        : TriviaGroupRegistry.GetSession(Config.TriviaGroup);

    /// <summary>Sends a trivia round message to this bot's own channel, and to every other bot sharing its TriviaGroup, so linked channels all see the game. A dead/disconnected peer is logged and skipped rather than failing the whole broadcast.</summary>
    private async Task BroadcastTriviaMessageAsync(string text)
    {
        await SendChatCommandAsync(text).ConfigureAwait(false);

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

    private async Task HandleTriviaCommandAsync(string rest, string username, Func<string, Task> reply)
    {
        if (!Config.TriviaFeatureEnabled)
        {
            await reply("Trivia isn't enabled for this bot.").ConfigureAwait(false);
            return;
        }

        var trimmed = rest.Trim();
        switch (trimmed.ToLowerInvariant())
        {
            case "on":
                await HandleTriviaOnAsync(reply, category: null).ConfigureAwait(false);
                break;

            case "off":
                _trivia.Stop();
                _triviaRoundCts?.Cancel();
                await reply("Trivia turned off.").ConfigureAwait(false);
                break;

            case "score":
                await HandleTriviaScoreAsync(reply).ConfigureAwait(false);
                break;

            case "join":
                await HandleTriviaJoinAsync(reply, username).ConfigureAwait(false);
                break;

            case "categories":
                await HandleTriviaCategoriesAsync(reply).ConfigureAwait(false);
                break;

            case "":
                await reply(_trivia.IsEnabled
                        ? "Use: !trivia ( off | score | categories )"
                        : "Use: !trivia ( on | <category> | score | categories )")
                    .ConfigureAwait(false);
                break;

            // Anything else is treated as an attempted category name, so "!trivia Blizzard"
            // starts a round using only that category's questions — see HandleTriviaOnAsync.
            default:
                await HandleTriviaOnAsync(reply, category: trimmed).ConfigureAwait(false);
                break;
        }
    }

    private async Task HandleTriviaOnAsync(Func<string, Task> reply, string? category)
    {
        if (_trivia.IsEnabled)
        {
            await reply("Trivia is already running.").ConfigureAwait(false);
            return;
        }

        var errors = new List<string>();
        var questions = TriviaBank.LoadAll(errors.Add);
        if (errors.Count > 0)
        {
            LogInfo($"Trivia: skipped {errors.Count} unparsable line(s): {string.Join("; ", errors.Take(5))}");
        }

        if (!string.IsNullOrEmpty(category))
        {
            var filtered = questions.Where(q => q.Category.Equals(category, StringComparison.OrdinalIgnoreCase)).ToList();
            if (filtered.Count == 0)
            {
                var known = KnownCategories(questions);
                var suggestion = known.Count > 0 ? $" Known categories: {string.Join(", ", known)}." : "";
                await reply($"No questions found for category \"{category}\".{suggestion}").ConfigureAwait(false);
                return;
            }

            questions = filtered;
        }

        if (questions.Count == 0)
        {
            await reply("No trivia questions found — add some .txt files to the Trivia folder (Open Config Folder).")
                .ConfigureAwait(false);
            return;
        }

        _trivia.Start(questions);
        var categoryNote = string.IsNullOrEmpty(category) ? "" : $" ({category})";
        var startedMessage = $"Trivia started{categoryNote} with {questions.Count} questions! Just answer in chat to play.";
        if (string.IsNullOrWhiteSpace(Config.TriviaGroup))
        {
            await reply(startedMessage).ConfigureAwait(false);
        }
        else
        {
            await BroadcastTriviaMessageAsync(startedMessage).ConfigureAwait(false);
        }

        _triviaRoundCts?.Cancel();
        _triviaRoundCts = new CancellationTokenSource();
        _ = RunTriviaRoundAsync(_triviaRoundCts.Token);
    }

    private Task HandleTriviaCategoriesAsync(Func<string, Task> reply)
    {
        var known = KnownCategories(TriviaBank.LoadAll());
        return known.Count == 0
            ? reply("No trivia categories found — add some .txt files to the Trivia folder (Open Config Folder).")
            : reply($"Trivia categories: {string.Join(", ", known)}. Start one with !trivia <category>.");
    }

    private static List<string> KnownCategories(IEnumerable<TriviaQuestion> questions) =>
        questions.Select(q => q.Category).Where(c => c.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(c => c).ToList();

    private Task HandleTriviaScoreAsync(Func<string, Task> reply)
    {
        var leaders = ClanRosterStore.Members
            .Where(m => m.TriviaScore != 0)
            .OrderByDescending(m => m.TriviaScore)
            .Take(10)
            .ToList();

        if (leaders.Count == 0)
        {
            return reply("No trivia scores yet — answer a question in chat to get on the board.");
        }

        var text = "Trivia Leader Board: " + string.Join(" ", leaders.Select(m => $"{m.Name}({m.TriviaScore})"));
        return reply(text);
    }

    /// <summary>
    /// Not required to play — anyone not banned can just answer in chat and
    /// score, same as BNU`Bot. This just proactively creates a (bare,
    /// unranked) roster entry for someone who wants to confirm they're
    /// tracked before they've answered anything, e.g. to show up on the
    /// leaderboard at 0. A no-op if they're already tracked.
    /// </summary>
    private Task HandleTriviaJoinAsync(Func<string, Task> reply, string username)
    {
        if (ClanRosterStore.Find(username) is not null)
        {
            return reply("You're already tracked.");
        }

        ClanRosterStore.Members.Add(new ClanMember { Name = username });
        ClanRosterStore.Save();
        return reply("You're in! Answer trivia questions in chat to earn points.");
    }

    private async Task RunTriviaRoundAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested && _trivia.IsEnabled)
        {
            var question = _trivia.AskNext();
            if (question is null)
            {
                await BroadcastTriviaMessageAsync("There are no trivia questions left; game over.").ConfigureAwait(false);
                _trivia.Stop();
                break;
            }

            var categoryText = question.Category.Length > 0 ? $" - Category: {question.Category}" : "";
            await BroadcastTriviaMessageAsync($"/me{categoryText} - Question: {question.QuestionText} - Hint: {question.Hint0}")
                .ConfigureAwait(false);

            _trivia.PendingAnswer = null;
            var (result, winner) = await WaitForAnswerOrTimeoutAsync(question, ct).ConfigureAwait(false);

            if (result == TriviaWaitResult.Cancelled)
            {
                // "!trivia off" (or 10-unanswered-streak auto-shutoff, handled
                // below) fired mid-question — stop silently, no "Time's up!"
                // or further messages for a round the operator just ended.
                break;
            }

            if (result == TriviaWaitResult.Answered && winner is { } win)
            {
                _trivia.RecordAnswered();
                await AnnounceWinnerAsync(question, win).ConfigureAwait(false);
            }
            else
            {
                _trivia.RecordTimeout();
                var correct = string.Join(", ", question.Answers.Select(a => $"\"{a}\""));
                await BroadcastTriviaMessageAsync($"/me - Time's up! The correct answer was {correct}").ConfigureAwait(false);

                if (_trivia.UnansweredStreak == 9)
                {
                    await BroadcastTriviaMessageAsync(
                        "Trivia will automatically shut off after the next question. To keep going, type !trivia on again once it stops.")
                        .ConfigureAwait(false);
                }

                if (_trivia.UnansweredStreak >= 10)
                {
                    await BroadcastTriviaMessageAsync("Auto-disabling trivia (10 unanswered questions in a row).").ConfigureAwait(false);
                    _trivia.Stop();
                    break;
                }
            }

            try
            {
                await Task.Delay(1000, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                break;
            }
        }
    }

    private async Task AnnounceWinnerAsync(TriviaQuestion question, (string Username, string MatchedAnswer) win)
    {
        var extra = "!";
        var member = ClanRosterStore.Find(win.Username);
        if (member is not null)
        {
            member.TriviaScore++;
            ClanRosterStore.Save();
            extra = $"! Your score is {member.TriviaScore}.";
        }

        var alternates = question.Answers.Where(a => a != win.MatchedAnswer).ToList();
        if (alternates.Count > 0)
        {
            extra += " Other acceptable answers were: " + string.Join(", ", alternates.Select(a => $"\"{a}\""));
        }

        await BroadcastTriviaMessageAsync($"/me - \"{win.MatchedAnswer}\" is correct, {win.Username}{extra}").ConfigureAwait(false);
    }

    private enum TriviaWaitResult
    {
        Answered,
        TimedOut,
        Cancelled,
    }

    /// <summary>
    /// Polls every 200ms for a matching answer (set by HandleChatEvent as
    /// chat arrives), revealing hints at +10s/+20s and giving up at 30s —
    /// mirrors TriviaEventHandler.triviaLoop's own poll loop rather than an
    /// event-driven wait, since that's the actual structure being ported.
    /// Re-checks _trivia.IsEnabled (not just the token) at the top of every
    /// iteration, so "!trivia off" mid-question stops within one tick instead
    /// of letting an already-due hint or timeout slip out afterward.
    /// </summary>
    private async Task<(TriviaWaitResult Result, (string Username, string MatchedAnswer)? Winner)> WaitForAnswerOrTimeoutAsync(
        TriviaQuestion question, CancellationToken ct)
    {
        var askedAt = DateTime.UtcNow;
        var hint1Given = false;
        var hint2Given = false;

        while (true)
        {
            if (!_trivia.IsEnabled || ct.IsCancellationRequested)
            {
                return (TriviaWaitResult.Cancelled, null);
            }

            if (_trivia.PendingAnswer is { } answer)
            {
                _trivia.PendingAnswer = null;
                return (TriviaWaitResult.Answered, answer);
            }

            var elapsed = DateTime.UtcNow - askedAt;

            if (!hint1Given && elapsed.TotalSeconds > 10)
            {
                hint1Given = true;
                await BroadcastTriviaMessageAsync($"/me - 20 seconds left! Hint: {question.Hint1}").ConfigureAwait(false);
            }

            if (!hint2Given && elapsed.TotalSeconds > 20)
            {
                hint2Given = true;
                await BroadcastTriviaMessageAsync($"/me - 10 seconds left! Hint: {question.Hint2}").ConfigureAwait(false);
            }

            if (elapsed.TotalSeconds >= 30)
            {
                return (TriviaWaitResult.TimedOut, null);
            }

            try
            {
                await Task.Delay(200, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                return (TriviaWaitResult.Cancelled, null);
            }
        }
    }
}
