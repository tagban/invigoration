using System.Reflection;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;
using Invigoration.Core.Trivia;

namespace Invigoration.Core.Tests;

/// <summary>
/// Covers the bug fixed alongside multi-channel SC2 chat: before this, an
/// answer typed in one joined channel could resolve a trivia question the
/// bot posed in a different one, since HandleChatEvent's TryMatchAnswer
/// check had no channel awareness at all.
/// </summary>
public class BotEngineTriviaChannelIsolationTests
{
    private static object GetTrivia(BotEngine engine) =>
        typeof(BotEngine).GetProperty("_trivia", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(engine)!;

    private static (string Username, string MatchedAnswer, string Source)? GetPendingAnswer(BotEngine engine) =>
        (ValueTuple<string, string, string>?)GetTrivia(engine).GetType().GetProperty("PendingAnswer")!.GetValue(GetTrivia(engine));

    private static Task InvokeHandleChatEvent(BotEngine engine, ChatEvent chatEvent)
    {
        var method = typeof(BotEngine).GetMethod("HandleChatEvent", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (Task)method.Invoke(engine, [chatEvent])!;
    }

    private static async Task<BotEngine> NewEngineWithRunningTriviaAsync(byte? triviaChannelIndex)
    {
        var config = new BotConfig { TriviaFeatureEnabled = true };
        var engine = new BotEngine(config);
        var trivia = GetTrivia(engine);
        var question = TriviaQuestion.Parse("What year was Diablo released?*1996", "Diablo");
        trivia.GetType().GetMethod("Start")!.Invoke(trivia, [new[] { question }]);
        trivia.GetType().GetMethod("AskNext")!.Invoke(trivia, []);
        typeof(BotEngine).GetField("_sc2TriviaChannelIndex", BindingFlags.NonPublic | BindingFlags.Instance)!
            .SetValue(engine, triviaChannelIndex);
        await Task.CompletedTask;
        return engine;
    }

    [Fact]
    public async Task AnswerFromWrongChannel_DoesNotResolveThePendingQuestion()
    {
        await using var engine = await NewEngineWithRunningTriviaAsync(triviaChannelIndex: 1);

        await InvokeHandleChatEvent(engine, new ChatEvent(ChatEventType.Talk, "player", 0, 0, "1996", ChannelIndex: 2));

        Assert.Null(GetPendingAnswer(engine));
    }

    [Fact]
    public async Task AnswerFromTheRightChannel_ResolvesThePendingQuestion()
    {
        await using var engine = await NewEngineWithRunningTriviaAsync(triviaChannelIndex: 1);

        await InvokeHandleChatEvent(engine, new ChatEvent(ChatEventType.Talk, "player", 0, 0, "1996", ChannelIndex: 1));

        var pending = GetPendingAnswer(engine);
        Assert.NotNull(pending);
        Assert.Equal("player", pending!.Value.Username);
    }

    [Fact]
    public async Task AnswerWithNoChannelIndex_AlwaysResolves()
    {
        // Classic BNCS/Chat-Telnet: ChannelIndex is always null (single-channel by
        // protocol), so the gate must never block a real answer there.
        await using var engine = await NewEngineWithRunningTriviaAsync(triviaChannelIndex: null);

        await InvokeHandleChatEvent(engine, new ChatEvent(ChatEventType.Talk, "player", 0, 0, "1996"));

        Assert.NotNull(GetPendingAnswer(engine));
    }
}
