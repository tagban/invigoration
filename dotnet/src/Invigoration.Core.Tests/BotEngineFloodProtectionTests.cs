using System.Diagnostics;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>
/// The flood-protection gate is process-wide (shared across every BotEngine,
/// not per-instance — see BotEngine.ChatSendGate's remarks for why), so
/// these tests only assert lower bounds, never "this is the very first send
/// so it should be instant" — that assumption doesn't hold once other tests
/// (or in production, other bots) may have already primed the shared gate.
/// </summary>
public class BotEngineFloodProtectionTests
{
    [Fact]
    public async Task SendChatCommandAsync_EnforcesMinimumDelayBetweenSends()
    {
        // No live connection needed: FramedTcpClient.SendAsync no-ops when
        // unconnected, but the flood-protection delay in SendChatCommandAsync
        // still runs before that no-op, so timing is testable in isolation.
        var config = new BotConfig { FloodProtectionDelayMs = 300 };
        await using var engine = new BotEngine(config);

        // Prime the shared gate so the delta below is caused by THIS test's
        // own delay, not leftover state from another test that ran first.
        await engine.SendChatCommandAsync("priming");

        var stopwatch = Stopwatch.StartNew();
        await engine.SendChatCommandAsync("second");
        stopwatch.Stop();

        Assert.True(
            stopwatch.ElapsedMilliseconds >= 250,
            $"Expected the second send to wait for most of the 300ms flood-protection delay; only {stopwatch.ElapsedMilliseconds}ms elapsed.");
    }

    [Fact]
    public async Task SendChatCommandAsync_TwoDifferentEngines_ShareTheSameThrottle()
    {
        // This is the exact scenario that got a test account flood-banned on
        // a second server: two linked bots each individually well-behaved,
        // but sending around the same moment. The gate must be shared, not
        // per-connection, for this to be caught.
        var configA = new BotConfig { FloodProtectionDelayMs = 300 };
        var configB = new BotConfig { FloodProtectionDelayMs = 300 };
        await using var engineA = new BotEngine(configA);
        await using var engineB = new BotEngine(configB);

        await engineA.SendChatCommandAsync("priming");

        var stopwatch = Stopwatch.StartNew();
        await engineB.SendChatCommandAsync("from a different engine, moments later");
        stopwatch.Stop();

        Assert.True(
            stopwatch.ElapsedMilliseconds >= 250,
            $"Expected engine B's send to be throttled by engine A's recent send; only {stopwatch.ElapsedMilliseconds}ms elapsed.");
    }

    [Fact]
    public async Task SendChatCommandAsync_ConcurrentCalls_AreSerializedThroughTheDelay()
    {
        var config = new BotConfig { FloodProtectionDelayMs = 200 };
        await using var engine = new BotEngine(config);

        await engine.SendChatCommandAsync("priming");

        var stopwatch = Stopwatch.StartNew();
        await Task.WhenAll(
            engine.SendChatCommandAsync("a"),
            engine.SendChatCommandAsync("b"),
            engine.SendChatCommandAsync("c"));
        stopwatch.Stop();

        // Three sends after priming, each at least 200ms apart: ~600ms minimum total.
        Assert.True(
            stopwatch.ElapsedMilliseconds >= 500,
            $"Expected concurrent sends to queue through the delay rather than firing together; only {stopwatch.ElapsedMilliseconds}ms elapsed.");
    }
}
