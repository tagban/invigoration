namespace Invigoration.Core.Trivia;

/// <summary>
/// Pure game-state for one bot's trivia round: which questions remain and
/// the current one. Anyone (not banned — that check lives in BotEngine,
/// which has the roster/config this class deliberately doesn't) can answer
/// without a separate opt-in step, matching BNU`Bot's own behavior. No
/// timers or chat I/O here — BotEngine.Trivia.cs owns the real-time
/// scheduling (hint reveals, timeouts) and sending replies, so this stays
/// simple to unit test.
/// </summary>
public sealed class TriviaSession
{
    private readonly List<TriviaQuestion> _pool = [];

    public bool IsEnabled { get; private set; }

    public TriviaQuestion? Current { get; private set; }

    /// <summary>
    /// Set by whichever engine's chat handler sees a matching answer come in
    /// and polled by whichever engine is running the round loop — usually
    /// the same engine, but not necessarily: when several bots share this
    /// session via BotConfig.TriviaGroup (e.g. one on Warcraft II, one on
    /// StarCraft II, same person's clan), the answer can arrive on any of
    /// their channels while only one of them owns the round loop. Source
    /// describes where the answer came from (server + chat room, or
    /// "Discord") — see BotEngine.Bncs.cs's DescribeChatSource — so a
    /// TriviaGroup round spanning several bots/channels can say which one a
    /// winner actually answered from.
    /// </summary>
    public (string Username, string MatchedAnswer, string Source)? PendingAnswer { get; set; }

    /// <summary>Consecutive questions nobody answered — trivia auto-disables at 10, matching BNU`Bot.</summary>
    public int UnansweredStreak { get; private set; }

    public int QuestionsRemaining => _pool.Count;

    public void Start(IEnumerable<TriviaQuestion> questions)
    {
        _pool.Clear();
        _pool.AddRange(questions);
        UnansweredStreak = 0;
        Current = null;
        IsEnabled = true;
    }

    public void Stop()
    {
        IsEnabled = false;
        Current = null;
    }

    /// <summary>Pulls a random remaining question as the new current one, or null if the pool is exhausted (round over).</summary>
    public TriviaQuestion? AskNext()
    {
        if (_pool.Count == 0)
        {
            Current = null;
            return null;
        }

        var index = Random.Shared.Next(_pool.Count);
        Current = _pool[index];
        _pool.RemoveAt(index);
        return Current;
    }

    /// <summary>Checks chat text against the current question, if there is one.</summary>
    public bool TryMatchAnswer(string text, out string matchedAnswer)
    {
        matchedAnswer = "";
        return Current is not null && Current.TryMatchAnswer(text, out matchedAnswer);
    }

    public void RecordAnswered() => UnansweredStreak = 0;

    public void RecordTimeout() => UnansweredStreak++;
}
