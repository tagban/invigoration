using System.Reflection;
using Invigoration.Core.Clan;
using Invigoration.Core.Config;
using Invigoration.Core.Trivia;

namespace Invigoration.Core.Tests;

/// <summary>In the shared "ClanRosterStore" xUnit collection since these touch the static roster directly.</summary>
[Collection("ClanRosterStore")]
public class BotEngineTriviaScoringTests
{
    private static double InvokePointsForStage(BotEngine engine, int stage)
    {
        var method = typeof(BotEngine).GetMethod("PointsForStage", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (double)method.Invoke(engine, [stage])!;
    }

    private static Task InvokeAnnounceWinnerAsync(BotEngine engine, TriviaQuestion question, (string Username, string MatchedAnswer, string Source) win, int stage)
    {
        var method = typeof(BotEngine).GetMethod("AnnounceWinnerAsync", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (Task)method.Invoke(engine, [question, win, stage])!;
    }

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

        Assert.Equal(1.25, InvokePointsForStage(engine, 0));
        Assert.Equal(1.0, InvokePointsForStage(engine, 1));
        Assert.Equal(0.75, InvokePointsForStage(engine, 2));
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

        var name = $"test-{Guid.NewGuid():N}";
        ClanRosterStore.Members.Add(new ClanMember { Name = name });
        try
        {
            var question = TriviaQuestion.Parse("What year was Diablo released?*1996", "Diablo");

            await InvokeAnnounceWinnerAsync(engine, question, (name, "1996", "test"), stage: 0);
            Assert.Equal(1.25, ClanRosterStore.Find(name)!.TriviaScore);

            await InvokeAnnounceWinnerAsync(engine, question, (name, "1996", "test"), stage: 1);
            Assert.Equal(2.25, ClanRosterStore.Find(name)!.TriviaScore);

            await InvokeAnnounceWinnerAsync(engine, question, (name, "1996", "test"), stage: 2);
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
            await InvokeAnnounceWinnerAsync(engine, question, (name, "Demon", "test"), stage: 0);

            Assert.Equal(1.25, ClanRosterStore.Find(name)!.TriviaScore);
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == name);
        }
    }
}
