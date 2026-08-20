using System.Reflection;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>
/// Confirms SID_REQUIREDWORK (Blizzard's ExtraWork anti-bot check) is
/// surfaced to the operator rather than silently ignored, and — just as
/// importantly — that handling it does nothing beyond logging. This project
/// deliberately never implements ExtraWork compliance (see BncsPacketId.
/// SID_REQUIREDWORK and BotEngine.Bncs.cs's HandleRequiredWork remarks).
/// </summary>
public class BotEngineRequiredWorkTests
{
    [Fact]
    public async Task HandleRequiredWork_LogsAnHonestWarning()
    {
        var config = new BotConfig();
        await using var engine = new BotEngine(config);
        var logged = new List<string>();
        engine.Log += segments => logged.Add(string.Concat(segments.Select(s => s.Text)));

        var method = typeof(BotEngine).GetMethod("HandleRequiredWork", BindingFlags.NonPublic | BindingFlags.Instance)!;
        await (Task)method.Invoke(engine, null)!;

        Assert.Contains(logged, l => l.Contains("anti-bot", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(logged, l => l.Contains("not implemented", StringComparison.OrdinalIgnoreCase));
    }
}
