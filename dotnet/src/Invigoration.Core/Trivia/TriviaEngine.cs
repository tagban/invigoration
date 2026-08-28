namespace Invigoration.Core.Trivia;

/// <summary>
/// The actual trivia round-runner — command handling, hint-reveal timing (initial hint with the
/// question, more at +10s/+20s, 30s total per question), scoring, and auto-shutoff after 10
/// consecutive unanswered questions — lifted out of BotEngine.Trivia.cs so the identical logic
/// runs for both Battle.net bots and Hotline sessions instead of two copies that inevitably drift.
/// Everything protocol-specific goes through <see cref="ITriviaHost"/>; this class never touches
/// BNCS or Hotline directly.
///
/// One long-lived instance per host (it owns the round's own CancellationTokenSource, which must
/// survive across the "!trivia on" and later "!trivia off" calls that start/stop it), but the
/// actual <see cref="TriviaSession"/> to operate on is resolved via <paramref name="resolveSession"/>
/// rather than fixed at construction — mirrors BotEngine's own "_trivia" property, which
/// re-resolves fresh every access rather than caching, so a session named via
/// BotConfig.TriviaGroup never goes stale after Config is edited/replaced. Resolved once per
/// logical operation (one HandleCommandAsync call, or once for a whole round in RunRoundAsync) —
/// not on every individual property touch — since the round's own state lives on whichever session
/// object it started against anyway.
/// </summary>
public sealed class TriviaEngine(ITriviaHost host, Func<TriviaSession> resolveSession)
{
    private CancellationTokenSource? _roundCts;

    public async Task HandleCommandAsync(string rest, Func<string, Task> reply)
    {
        if (!host.TriviaFeatureEnabled)
        {
            await reply("Trivia isn't enabled for this bot.").ConfigureAwait(false);
            return;
        }

        var session = resolveSession();
        var trimmed = rest.Trim();
        switch (trimmed.ToLowerInvariant())
        {
            case "on":
                await HandleOnAsync(session, reply, category: null, gameshowMode: false).ConfigureAwait(false);
                break;

            case "all":
                await HandleOnAsync(session, reply, category: null, gameshowMode: true).ConfigureAwait(false);
                break;

            case "off":
                session.Stop();
                _roundCts?.Cancel();
                await reply("Trivia turned off.").ConfigureAwait(false);
                break;

            case "score":
                await HandleScoreAsync(reply).ConfigureAwait(false);
                break;

            case "categories":
                await HandleCategoriesAsync(reply).ConfigureAwait(false);
                break;

            case "":
                await reply(session.IsEnabled
                        ? "Use: !trivia ( off | score | categories )"
                        : "Use: !trivia ( on | all | <category> | score | categories )")
                    .ConfigureAwait(false);
                break;

            // Anything else is treated as an attempted category name, so "!trivia Blizzard"
            // starts a round using only that category's questions — see HandleOnAsync.
            default:
                await HandleOnAsync(session, reply, category: trimmed, gameshowMode: false).ConfigureAwait(false);
                break;
        }
    }

    private async Task HandleOnAsync(TriviaSession session, Func<string, Task> reply, string? category, bool gameshowMode)
    {
        if (session.IsEnabled)
        {
            await reply("Trivia is already running.").ConfigureAwait(false);
            return;
        }

        var errors = new List<string>();
        var questions = TriviaBank.LoadAll(errors.Add);
        if (errors.Count > 0)
        {
            host.LogParseErrors(errors);
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

        session.Start(questions);
        host.OnRoundStarting();
        var categoryNote = string.IsNullOrEmpty(category) ? "" : $" ({category})";
        var startedMessage = $"Trivia started{categoryNote} with {questions.Count} questions! Just answer in chat to play.";
        await host.AnnounceStartedAsync(startedMessage, reply).ConfigureAwait(false);

        _roundCts?.Cancel();
        _roundCts = new CancellationTokenSource();
        _ = RunRoundAsync(session, gameshowMode, _roundCts.Token);
    }

    private Task HandleCategoriesAsync(Func<string, Task> reply)
    {
        var known = KnownCategories(TriviaBank.LoadAll());
        return known.Count == 0
            ? reply("No trivia categories found — add some .txt files to the Trivia folder (Open Config Folder).")
            : reply($"Trivia categories: {string.Join(", ", known)}. Start one with !trivia <category>.");
    }

    private static List<string> KnownCategories(IEnumerable<TriviaQuestion> questions) =>
        questions.Select(q => q.Category).Where(c => c.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(c => c).ToList();

    private Task HandleScoreAsync(Func<string, Task> reply)
    {
        var body = host.FormatLeaderboard();
        return reply(body.Length == 0
            ? "No trivia scores yet — answer a question in chat to get on the board."
            : $"Trivia Leader Board: {body}");
    }

    /// <summary>gameshowMode ("!trivia all") announces each question's category as its own message, with a short pause, before the question itself — versus the default single combined "- Category: X - Question: ..." line.</summary>
    private async Task RunRoundAsync(TriviaSession session, bool gameshowMode, CancellationToken ct)
    {
        while (!ct.IsCancellationRequested && session.IsEnabled)
        {
            var question = session.AskNext();
            if (question is null)
            {
                await host.BroadcastAsync("There are no trivia questions left; game over.").ConfigureAwait(false);
                session.Stop();
                break;
            }

            if (gameshowMode && question.Category.Length > 0)
            {
                await host.BroadcastAsync($"/me \U0001F3AF Category: {question.Category}!").ConfigureAwait(false);
                try
                {
                    await Task.Delay(1200, ct).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    break;
                }
            }

            // In gameshow mode the category was just announced as its own message above, so it's
            // left out of the question line to avoid saying it twice.
            var categoryText = !gameshowMode && question.Category.Length > 0 ? $" - Category: {question.Category}" : "";
            var questionMessage = question.IsMultipleChoice
                ? $"/me{categoryText} - Question: {question.QuestionText} - {string.Join("  ", question.Choices.Select(c => $"{c.Letter}) {c.Text}"))}"
                : $"/me{categoryText} - Question: {question.QuestionText} - Hint: {question.Hint0}";
            await host.BroadcastAsync(questionMessage).ConfigureAwait(false);

            session.PendingAnswer = null;
            var (result, winner, stage) = await WaitForAnswerOrTimeoutAsync(session, question, ct).ConfigureAwait(false);

            if (result == TriviaWaitResult.Cancelled)
            {
                // "!trivia off" (or 10-unanswered-streak auto-shutoff, handled below) fired
                // mid-question — stop silently, no "Time's up!" or further messages for a round
                // the operator just ended.
                break;
            }

            if (result == TriviaWaitResult.Answered && winner is { } win)
            {
                session.RecordAnswered();
                await AnnounceWinnerAsync(question, win, stage).ConfigureAwait(false);
            }
            else
            {
                session.RecordTimeout();
                // For multiple choice, Answers[0] is always the correct option's actual text
                // (Answers[1] is just its letter, not a second real synonym worth repeating here)
                // — see TriviaQuestion's private multiple-choice constructor.
                var correct = question.IsMultipleChoice
                    ? $"\"{question.Answers[0]}\""
                    : string.Join(", ", question.Answers.Select(a => $"\"{a}\""));
                await host.BroadcastAsync($"/me - Time's up! The correct answer was {correct}").ConfigureAwait(false);

                if (session.UnansweredStreak == 9)
                {
                    await host.BroadcastAsync(
                        "Trivia will automatically shut off after the next question. To keep going, type !trivia on again once it stops.")
                        .ConfigureAwait(false);
                }

                if (session.UnansweredStreak >= 10)
                {
                    await host.BroadcastAsync("Auto-disabling trivia (10 unanswered questions in a row).").ConfigureAwait(false);
                    session.Stop();
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

    /// <summary>Which graduated-scoring tier an answer landed in — 0 = before the first hint, 1 = after the first hint, 2 = after the second hint. Multiple-choice questions never advance past 0, since there's no hint progression to speak of.</summary>
    private double PointsForStage(int stage) => stage switch
    {
        0 => host.PointsBeforeFirstHint,
        1 => host.PointsAfterFirstHint,
        _ => host.PointsAfterSecondHint,
    };

    private async Task AnnounceWinnerAsync(TriviaQuestion question, (string Username, string MatchedAnswer, string Source) win, int stage)
    {
        var points = PointsForStage(stage);
        var scoreFragment = host.RecordScore(win.Username, points);
        var extra = "!" + scoreFragment;

        // The multiple-choice Answers list carries the correct text plus its assigned letter as
        // two equally-valid ways to match — not real alternate synonyms worth listing back to the
        // winner.
        var alternates = question.IsMultipleChoice
            ? []
            : question.Answers.Where(a => a != win.MatchedAnswer).ToList();
        if (alternates.Count > 0)
        {
            extra += " Other acceptable answers were: " + string.Join(", ", alternates.Select(a => $"\"{a}\""));
        }

        await host.BroadcastAsync($"/me - \"{win.MatchedAnswer}\" is correct, {win.Username} (from {win.Source}){extra}").ConfigureAwait(false);
    }

    private enum TriviaWaitResult
    {
        Answered,
        TimedOut,
        Cancelled,
    }

    /// <summary>
    /// Polls every 200ms for a matching answer (set on TriviaSession.PendingAnswer by whichever
    /// host's own chat handler sees it arrive), revealing hints at +10s/+20s and giving up at 30s.
    /// Re-checks session.IsEnabled (not just the token) at the top of every iteration, so
    /// "!trivia off" mid-question stops within one tick instead of letting an already-due hint or
    /// timeout slip out afterward. Multiple-choice questions skip both hint reveals entirely —
    /// every option is already visible, so the returned stage stays 0 (see PointsForStage) for the
    /// whole 30-second window.
    /// </summary>
    private async Task<(TriviaWaitResult Result, (string Username, string MatchedAnswer, string Source)? Winner, int Stage)> WaitForAnswerOrTimeoutAsync(
        TriviaSession session, TriviaQuestion question, CancellationToken ct)
    {
        var askedAt = DateTime.UtcNow;
        var hint1Given = false;
        var hint2Given = false;
        var stage = 0;

        while (true)
        {
            if (!session.IsEnabled || ct.IsCancellationRequested)
            {
                return (TriviaWaitResult.Cancelled, null, 0);
            }

            if (session.PendingAnswer is { } answer)
            {
                session.PendingAnswer = null;
                return (TriviaWaitResult.Answered, answer, stage);
            }

            var elapsed = DateTime.UtcNow - askedAt;

            if (!question.IsMultipleChoice)
            {
                if (!hint1Given && elapsed.TotalSeconds > 10)
                {
                    hint1Given = true;
                    stage = 1;
                    await host.BroadcastAsync($"/me - 20 seconds left! Hint: {question.Hint1}").ConfigureAwait(false);
                }

                if (!hint2Given && elapsed.TotalSeconds > 20)
                {
                    hint2Given = true;
                    stage = 2;
                    await host.BroadcastAsync($"/me - 10 seconds left! Hint: {question.Hint2}").ConfigureAwait(false);
                }
            }

            if (elapsed.TotalSeconds >= 30)
            {
                return (TriviaWaitResult.TimedOut, null, 0);
            }

            try
            {
                await Task.Delay(200, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                return (TriviaWaitResult.Cancelled, null, 0);
            }
        }
    }
}
