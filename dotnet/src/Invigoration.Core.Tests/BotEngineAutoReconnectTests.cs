using System.Reflection;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>
/// MaybeScheduleAutoReconnect is private and fires a background task, so
/// these check the decision it makes (does it log "reconnecting", which
/// happens before the actual delay/connect attempt) rather than driving a
/// real reconnect, which would need network access.
/// </summary>
public class BotEngineAutoReconnectTests
{
    private static void InvokeMaybeScheduleAutoReconnect(BotEngine engine)
    {
        var method = typeof(BotEngine).GetMethod("MaybeScheduleAutoReconnect", BindingFlags.NonPublic | BindingFlags.Instance)!;
        method.Invoke(engine, null);
    }

    [Fact]
    public async Task MaybeScheduleAutoReconnect_FeatureDisabled_DoesNotSchedule()
    {
        var config = new BotConfig { AutoReconnect = false };
        await using var engine = new BotEngine(config);
        var logged = new List<string>();
        engine.Log += segments => logged.Add(string.Concat(segments.Select(s => s.Text)));

        InvokeMaybeScheduleAutoReconnect(engine);
        await Task.Delay(100);

        Assert.DoesNotContain(logged, l => l.Contains("reconnecting", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task MaybeScheduleAutoReconnect_EnabledAfterUnexpectedDisconnect_SchedulesReconnect()
    {
        var config = new BotConfig { AutoReconnect = true, AutoReconnectDelaySeconds = 9999 };
        await using var engine = new BotEngine(config);
        var logged = new List<string>();
        engine.Log += segments => logged.Add(string.Concat(segments.Select(s => s.Text)));

        InvokeMaybeScheduleAutoReconnect(engine);
        await Task.Delay(100);

        Assert.Contains(logged, l => l.Contains("reconnecting", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task MaybeScheduleAutoReconnect_AfterIntentionalDisconnect_DoesNotSchedule()
    {
        var config = new BotConfig { AutoReconnect = true, AutoReconnectDelaySeconds = 9999 };
        await using var engine = new BotEngine(config);
        var logged = new List<string>();
        engine.Log += segments => logged.Add(string.Concat(segments.Select(s => s.Text)));

        // DisconnectAsync marks the next disconnect as intentional, so a
        // subsequent unsolicited call to the scheduler (as would come from
        // the underlying socket's own Disconnected event firing after Close())
        // must not trigger a reconnect.
        await engine.DisconnectAsync();
        InvokeMaybeScheduleAutoReconnect(engine);
        await Task.Delay(100);

        Assert.DoesNotContain(logged, l => l.Contains("reconnecting", StringComparison.OrdinalIgnoreCase));
    }
}

/// <summary>
/// Regression: a bot with auto-reconnect on was being disconnected by its own client every
/// AutoReconnectDelaySeconds after logging on, forever (seen live: 35s cycles, server log "client
/// closed the connection"). A reconnect that fired after the bot was already back on replaced the
/// live connection, that close was treated as an unexpected drop, and the drop scheduled the next
/// reconnect.
/// </summary>
public class BotEngineReconnectLoopTests
{
    private const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Instance;

    private static List<string> Watch(BotEngine engine)
    {
        var logged = new List<string>();
        engine.Log += segments => { lock (logged) { logged.Add(string.Concat(segments.Select(s => s.Text))); } };
        return logged;
    }

    private static void RaiseBncsDisconnected(BotEngine engine)
    {
        var bncs = typeof(BotEngine).GetField("_bncs", Private)!.GetValue(engine)!;
        var handler = (Action<Exception?>?)typeof(Invigoration.Core.Networking.FramedTcpClient).GetField("Disconnected", Private)!.GetValue(bncs);
        handler?.Invoke(null);
    }

    [Fact]
    public async Task ReplacingALiveConnection_IsNotADrop_AndSchedulesNoReconnect()
    {
        await using var engine = new BotEngine(new BotConfig { AutoReconnect = true, AutoReconnectDelaySeconds = 9999 });
        var logged = Watch(engine);
        typeof(BotEngine).GetField("_replacingBncsConnection", Private)!.SetValue(engine, 1);

        RaiseBncsDisconnected(engine);
        await Task.Delay(100);

        Assert.DoesNotContain(logged, l => l.Contains("reconnecting", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(logged, l => l.Contains("Battle.net disconnected", StringComparison.Ordinal));
        Assert.Equal(0, (int)typeof(BotEngine).GetField("_replacingBncsConnection", Private)!.GetValue(engine)!);
    }

    [Fact]
    public async Task ARealDrop_StillSchedulesTheReconnect()
    {
        await using var engine = new BotEngine(new BotConfig { AutoReconnect = true, AutoReconnectDelaySeconds = 9999 });
        var logged = Watch(engine);

        RaiseBncsDisconnected(engine);
        await Task.Delay(100);

        Assert.Contains(logged, l => l.Contains("reconnecting", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task LoggingOn_CancelsAReconnectStillCountingDown()
    {
        await using var engine = new BotEngine(new BotConfig { AutoReconnect = true, AutoReconnectDelaySeconds = 9999 });
        RaiseBncsDisconnected(engine);
        var pending = (CancellationTokenSource)typeof(BotEngine).GetField("_autoReconnectCts", Private)!.GetValue(engine)!;
        Assert.False(pending.IsCancellationRequested);

        await (Task)typeof(BotEngine).GetMethod("OnLoggedOnAsync", Private)!.Invoke(engine, [false, false])!;

        Assert.True(pending.IsCancellationRequested);
    }

    [Fact]
    public async Task AReconnectThatFiresWhileLoggedOn_DoesNothing()
    {
        await using var engine = BotEngineChatGateTests.MarkLoggedOn(new BotEngine(new BotConfig { AutoReconnect = true }));
        var logged = Watch(engine);

        await (Task)typeof(BotEngine).GetMethod("RunAutoReconnectAsync", Private)!.Invoke(engine, [1, CancellationToken.None])!;

        Assert.DoesNotContain(logged, l => l.Contains("connecting to", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(logged, l => l.Contains("reconnecting in 1s", StringComparison.Ordinal));
    }
}
