using System.Reflection;
using Invigoration.Core.Clan;
using Invigoration.Core.Config;
using Invigoration.Core.Trivia;

namespace Invigoration.Core.Tests;

/// <summary>
/// Exercises TriviaEngine's scoring logic (PointsForStage/AnnounceWinnerAsync — lifted out of
/// BotEngine.Trivia.cs into the shared Trivia.TriviaEngine, see its remarks) through a real
/// BotEngine acting as the ITriviaHost, so this still validates the full path including
/// BotEngine's own RecordScore -&gt; Clan.ClanRosterStore integration, not just TriviaEngine in
/// isolation. In the shared "ClanRosterStore" xUnit collection since these touch the static
/// roster directly.
/// </summary>
[Collection("ClanRosterStore")]
public class BotEngineTriviaScoringTests
{
    private static double InvokePointsForStage(TriviaEngine triviaEngine, int stage)
    {
        var method = typeof(TriviaEngine).GetMethod("PointsForStage", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (double)method.Invoke(triviaEngine, [stage])!;
    }

    private static Task InvokeAnnounceWinnerAsync(TriviaEngine triviaEngine, TriviaQuestion question, (string Username, string MatchedAnswer, string Source) win, int stage)
    {
        var method = typeof(TriviaEngine).GetMethod("AnnounceWinnerAsync", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (Task)method.Invoke(triviaEngine, [question, win, stage])!;
    }

    private static TriviaEngine CreateTriviaEngine(BotEngine engine) => new(engine, () => new TriviaSession());

    [Fact]
    public async Task PointsForStage_ReturnsConfiguredValuePerTier()
    {
        var config = new BotConfig
        {
            TriviaPointsBeforeFirstHint = 1.25,
            TriviaPointsAfterFirstHint = 1.0,
            TriviaPointsAfterSecondHint = 0.75,
        };
        await using var engine = new BotEngine(config);
        var triviaEngine = CreateTriviaEngine(engine);

        Assert.Equal(1.25, InvokePointsForStage(triviaEngine, 0));
        Assert.Equal(1.0, InvokePointsForStage(triviaEngine, 1));
        Assert.Equal(0.75, InvokePointsForStage(triviaEngine, 2));
    }

    [Fact]
    public async Task AnnounceWinnerAsync_AwardsTheRightTiersPoints()
    {
        var config = new BotConfig
        {
            TriviaPointsBeforeFirstHint = 1.25,
            TriviaPointsAfterFirstHint = 1.0,
            TriviaPointsAfterSecondHint = 0.75,
        };
        await using var engine = new BotEngine(config);
        var triviaEngine = CreateTriviaEngine(engine);

        var name = $"test-{Guid.NewGuid():N}";
        ClanRosterStore.Members.Add(new ClanMember { Name = name });
        try
        {
            var question = TriviaQuestion.Parse("What year was Diablo released?*1996", "Diablo");

            await InvokeAnnounceWinnerAsync(triviaEngine, question, (name, "1996", "test"), stage: 0);
            Assert.Equal(1.25, ClanRosterStore.Find(name)!.TriviaScore);

            await InvokeAnnounceWinnerAsync(triviaEngine, question, (name, "1996", "test"), stage: 1);
            Assert.Equal(2.25, ClanRosterStore.Find(name)!.TriviaScore);

            await InvokeAnnounceWinnerAsync(triviaEngine, question, (name, "1996", "test"), stage: 2);
            Assert.Equal(3.0, ClanRosterStore.Find(name)!.TriviaScore);
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == name);
        }
    }

    [Fact]
    public async Task AnnounceWinnerAsync_MultipleChoiceQuestion_AwardsFlatBeforeFirstHintPoints()
    {
        var config = new BotConfig { TriviaPointsBeforeFirstHint = 1.25 };
        await using var engine = new BotEngine(config);
        var triviaEngine = CreateTriviaEngine(engine);
        var name = $"test-{Guid.NewGuid():N}";
        ClanRosterStore.Members.Add(new ClanMember { Name = name });
        try
        {
            var question = TriviaQuestion.CreateMultipleChoice("Diablo", "What class of enemy is Diablo?", "Demon", ["Undead"]);

            // Multiple choice never advances past stage 0 (see
            // WaitForAnswerOrTimeoutAsync, which skips hint reveals for it) —
            // confirms AnnounceWinnerAsync doesn't crash or misbehave when
            // handed a multiple-choice question, and always uses the
            // "before first hint" tier for it regardless of real elapsed time.
            await InvokeAnnounceWinnerAsync(triviaEngine, question, (name, "Demon", "test"), stage: 0);

            Assert.Equal(1.25, ClanRosterStore.Find(name)!.TriviaScore);
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == name);
        }
    }
}
