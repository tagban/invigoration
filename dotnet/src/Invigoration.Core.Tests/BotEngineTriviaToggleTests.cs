using System.Reflection;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

public class BotEngineTriviaToggleTests
{
    private static Task InvokeHandleTriviaCommandAsync(BotEngine engine, string rest, string username, Func<string, Task> reply)
    {
        var method = typeof(BotEngine).GetMethod("HandleTriviaCommandAsync", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (Task)method.Invoke(engine, [rest, username, reply])!;
    }

    [Fact]
    public async Task HandleTriviaCommandAsync_FeatureDisabled_RepliesDisabledForEveryCommand()
    {
        var config = new BotConfig { TriviaFeatureEnabled = false };
        await using var engine = new BotEngine(config);
        var replies = new List<string>();
        Task Reply(string text)
        {
            replies.Add(text);
            return Task.CompletedTask;
        }

        await InvokeHandleTriviaCommandAsync(engine, "on", "SomeUser", Reply);
        await InvokeHandleTriviaCommandAsync(engine, "score", "SomeUser", Reply);

        Assert.Equal(2, replies.Count);
        Assert.All(replies, r => Assert.Equal("Trivia isn't enabled for this bot.", r));
    }

    [Fact]
    public async Task HandleTriviaCommandAsync_FeatureEnabled_ScoreCommandRespondsNormally()
    {
        var config = new BotConfig { TriviaFeatureEnabled = true };
        await using var engine = new BotEngine(config);
        var replies = new List<string>();
        Task Reply(string text)
        {
            replies.Add(text);
            return Task.CompletedTask;
        }

        await InvokeHandleTriviaCommandAsync(engine, "score", "SomeUser", Reply);

        Assert.Single(replies);
        Assert.NotEqual("Trivia isn't enabled for this bot.", replies[0]);
    }

    [Fact]
    public async Task HandleTriviaCommandAsync_Categories_ListsBundledCategories()
    {
        var config = new BotConfig { TriviaFeatureEnabled = true };
        await using var engine = new BotEngine(config);
        var replies = new List<string>();
        Task Reply(string text)
        {
            replies.Add(text);
            return Task.CompletedTask;
        }

        await InvokeHandleTriviaCommandAsync(engine, "categories", "SomeUser", Reply);

        Assert.Single(replies);
        Assert.Contains("Diablo", replies[0]);
        Assert.Contains("Blizzard", replies[0]);
        Assert.Contains("Music", replies[0]);
    }

    [Fact]
    public async Task HandleTriviaCommandAsync_UnknownCategory_DoesNotStartAndListsKnownCategories()
    {
        var config = new BotConfig { TriviaFeatureEnabled = true };
        await using var engine = new BotEngine(config);
        var replies = new List<string>();
        Task Reply(string text)
        {
            replies.Add(text);
            return Task.CompletedTask;
        }

        await InvokeHandleTriviaCommandAsync(engine, "NotARealCategory", "SomeUser", Reply);

        Assert.Single(replies);
        Assert.Contains("No questions found", replies[0]);
        Assert.Contains("Diablo", replies[0]);
    }

    [Fact]
    public async Task HandleTriviaCommandAsync_ValidCategory_StartsFilteredRound()
    {
        var config = new BotConfig { TriviaFeatureEnabled = true };
        await using var engine = new BotEngine(config);
        var replies = new List<string>();
        Task Reply(string text)
        {
            replies.Add(text);
            return Task.CompletedTask;
        }

        await InvokeHandleTriviaCommandAsync(engine, "Music", "SomeUser", Reply);

        Assert.Single(replies);
        Assert.Contains("Music", replies[0]);
        Assert.Contains("started", replies[0], StringComparison.OrdinalIgnoreCase);

        // Stop immediately so the background round loop this started doesn't keep running past the test.
        await InvokeHandleTriviaCommandAsync(engine, "off", "SomeUser", Reply);
    }
}
