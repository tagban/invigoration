using System.Reflection;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>
/// Confirms SID_REQUIREDWORK (historically ExtraWork) handling does nothing beyond a debug-level
/// log line — this project deliberately never implements ExtraWork compliance (see
/// BncsPacketId.SID_REQUIREDWORK and BotEngine.Bncs.cs's HandleRequiredWork remarks). Downgraded
/// from a always-visible LogWarning to LogDebug 2026-08-24: the user confirmed via a former
/// Blizzard employee that ExtraWork isn't actually enforced on live official Battle.net anymore,
/// so it's no longer worth surfacing as an alarming "anti-bot check" warning on every connection.
/// </summary>
public class BotEngineRequiredWorkTests
{
    [Fact]
    public async Task HandleRequiredWork_WithDebugModeOn_LogsAtDebugLevel()
    {
        var config = new BotConfig();
        await using var engine = new BotEngine(config) { DebugMode = true };
        var logged = new List<string>();
        engine.Log += segments => logged.Add(string.Concat(segments.Select(s => s.Text)));

        var method = typeof(BotEngine).GetMethod("HandleRequiredWork", BindingFlags.NonPublic | BindingFlags.Instance)!;
        await (Task)method.Invoke(engine, null)!;

        Assert.Contains(logged, l => l.Contains("SID_REQUIREDWORK", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(logged, l => l.Contains("not implemented", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task HandleRequiredWork_WithDebugModeOff_LogsNothing()
    {
        var config = new BotConfig();
        await using var engine = new BotEngine(config) { DebugMode = false };
        var logged = new List<string>();
        engine.Log += segments => logged.Add(string.Concat(segments.Select(s => s.Text)));

        var method = typeof(BotEngine).GetMethod("HandleRequiredWork", BindingFlags.NonPublic | BindingFlags.Instance)!;
        await (Task)method.Invoke(engine, null)!;

        Assert.Empty(logged);
    }
}
