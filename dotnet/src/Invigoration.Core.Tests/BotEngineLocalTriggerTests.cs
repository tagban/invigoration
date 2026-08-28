using System.Reflection;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>
/// Covers a deliberate behavior change: locally (typed in the bot's own
/// input box) only "/" runs a command now — the configured Trigger
/// character used to also work locally, but now only works when heard from
/// another user in the channel. Uses "fudd" (toggles _session.FuddMode) as
/// an observable side effect, since command replies go out over a
/// (disconnected, no-op) wire send rather than through a capturable event.
/// </summary>
public class BotEngineLocalTriggerTests
{
    private static bool GetFuddMode(BotEngine engine)
    {
        var field = typeof(BotEngine).GetField("_session", BindingFlags.NonPublic | BindingFlags.Instance)!;
        var session = field.GetValue(engine)!;
        return (bool)session.GetType().GetProperty("FuddMode")!.GetValue(session)!;
    }

    private static Task InvokeRemoteCommand(BotEngine engine, string username, string message)
    {
        var method = typeof(BotEngine).GetMethod("HandleCommandAsync", BindingFlags.NonPublic | BindingFlags.Instance,
            null, [typeof(string), typeof(string), typeof(bool), typeof(byte?)], null)!;
        return (Task)method.Invoke(engine, [username, message, false, null])!;
    }

    [Fact]
    public async Task RunLocalCommandAsync_TriggerPrefixed_IsNotRunAsCommand()
    {
        var config = new BotConfig { Trigger = "!" };
        await using var engine = new BotEngine(config);

        await engine.RunLocalCommandAsync("!fudd");

        Assert.False(GetFuddMode(engine));
    }

    [Fact]
    public async Task RunLocalCommandAsync_SlashPrefixed_IsRunAsCommand()
    {
        var config = new BotConfig { Trigger = "!" };
        await using var engine = new BotEngine(config);

        await engine.RunLocalCommandAsync("/fudd");

        Assert.True(GetFuddMode(engine));
    }

    [Fact]
    public async Task HandleCommandAsync_TriggerPrefixedFromBotMaster_StillRunsRemotely()
    {
        // "user" (sets _session.TargetUser), not "fudd" here — fudd is one of the commands
        // BotEngine.Commands.cs's LocalOnlyCommands blocks from ever running remotely as of
        // 2026-08-26 (regardless of who sends it, even the bot master), so it's no longer a valid
        // "does the trigger character still work remotely at all" probe.
        var config = new BotConfig { Trigger = "!", BotMaster = "TheMaster" };
        await using var engine = new BotEngine(config);

        await InvokeRemoteCommand(engine, "TheMaster", "!user SomeTarget");

        Assert.Equal("SomeTarget", GetTargetUser(engine));
    }

    private static string GetTargetUser(BotEngine engine)
    {
        var field = typeof(BotEngine).GetField("_session", BindingFlags.NonPublic | BindingFlags.Instance)!;
        var session = field.GetValue(engine)!;
        return (string)session.GetType().GetProperty("TargetUser")!.GetValue(session)!;
    }
}
