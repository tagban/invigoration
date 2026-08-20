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
