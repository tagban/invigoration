using System.Reflection;
using Invigoration.Core.Auth;
using Invigoration.Core.Chat;
using Invigoration.Core.Commands;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>
/// Classic BNCS chat waits until the server can accept it and is split to fit — BNETDocs' Atlas
/// closes the connection over chat sent before logon, plain chat sent outside a channel, or a
/// line past 223 characters. Uses unconnected engines: the socket send no-ops, and SelfChatSent
/// (the local echo of each plain line actually sent) shows what went out.
/// </summary>
public class BotEngineChatGateTests
{
    internal static BotEngine MarkLoggedOn(BotEngine engine)
    {
        var auth = (AuthState)typeof(BotEngine).GetField("_auth", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(engine)!;
        auth.LoggedOnToBncs = true;
        return engine;
    }

    internal static BotEngine MarkInChannel(BotEngine engine)
    {
        MarkLoggedOn(engine);
        var session = (BotSessionState)typeof(BotEngine).GetField("_session", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(engine)!;
        session.CurrentChannelName = "Town Square";
        return engine;
    }

    private static (List<string> Logs, List<string> Echoes) Watch(BotEngine engine)
    {
        var logs = new List<string>();
        var echoes = new List<string>();
        engine.Log += s => logs.Add(string.Concat(s.Select(x => x.Text)));
        engine.SelfChatSent += s => echoes.Add(string.Concat(s.Select(x => x.Text)));
        return (logs, echoes);
    }

    [Fact]
    public async Task BeforeLogon_NothingIsSent()
    {
        await using var engine = new BotEngine(new BotConfig { FloodProtectionDelayMs = 0 });
        var (logs, echoes) = Watch(engine);

        await engine.SendChatCommandAsync("hello");
        await engine.SendChatCommandAsync("/w Someone hi");

        Assert.Empty(echoes);
        Assert.Equal(2, logs.Count(l => l.Contains("Not sent (not logged on to Battle.net)", StringComparison.Ordinal)));
    }

    [Fact]
    public async Task LoggedOnButNotInAChannel_HoldsPlainChat_ButLetsCommandsThrough()
    {
        await using var engine = MarkLoggedOn(new BotEngine(new BotConfig { FloodProtectionDelayMs = 0 }));
        var (logs, echoes) = Watch(engine);

        await engine.SendChatCommandAsync("hello");
        await engine.SendChatCommandAsync("/w Someone hi");

        Assert.Empty(echoes);
        Assert.Single(logs, l => l.Contains("Not sent (not in a channel yet): hello", StringComparison.Ordinal));
        Assert.DoesNotContain(logs, l => l.Contains("/w Someone", StringComparison.Ordinal));
    }

    [Fact]
    public async Task InAChannel_PlainChatIsSent()
    {
        await using var engine = MarkInChannel(new BotEngine(new BotConfig { Username = "Tagban", FloodProtectionDelayMs = 0 }));
        var (logs, echoes) = Watch(engine);

        await engine.SendChatCommandAsync("hello");

        Assert.Single(echoes, e => e.EndsWith("hello", StringComparison.Ordinal));
        Assert.DoesNotContain(logs, l => l.StartsWith("Not sent", StringComparison.Ordinal));
    }

    [Fact]
    public async Task LongChat_GoesOutAsSeveralLinesWithinTheLimit()
    {
        await using var engine = MarkInChannel(new BotEngine(new BotConfig { Username = "Tagban", FloodProtectionDelayMs = 0 }));
        var (_, echoes) = Watch(engine);
        var text = string.Join(' ', Enumerable.Range(0, 80).Select(i => $"word{i:00}"));

        await engine.SendChatCommandAsync(text);

        Assert.True(echoes.Count > 1);
        var bodies = echoes.Select(e => e["Tagban: ".Length..]).ToList();
        Assert.All(bodies, b => Assert.InRange(b.Length, 1, ChatLineSplitter.MaxLineLength));
        Assert.Equal(text, string.Join(' ', bodies));
    }

    [Fact]
    public async Task EmptyText_IsIgnored()
    {
        await using var engine = MarkInChannel(new BotEngine(new BotConfig { FloodProtectionDelayMs = 0 }));
        var (logs, echoes) = Watch(engine);

        await engine.SendChatCommandAsync("");

        Assert.Empty(echoes);
        Assert.Empty(logs);
    }
}
