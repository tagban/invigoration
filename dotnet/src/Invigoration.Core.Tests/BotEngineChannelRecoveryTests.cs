using System.Reflection;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;
using Invigoration.Core.Protocol;

namespace Invigoration.Core.Tests;

/// <summary>
/// HandleChatEvent is private (no public entry point without simulating a
/// raw BNCS frame through the receive loop), so these call it via reflection
/// with a hand-built SID_CHATEVENT frame — same technique already used for
/// BotEngine.IsAuthorized in BotEngineAuthorizationTests.
/// </summary>
public class BotEngineChannelRecoveryTests
{
    private static byte[] BuildInfoEventFrame(string text) =>
        new PacketWriter()
            .WriteDword((uint)ChatEventType.Info)
            .WriteDword(0)
            .WriteDword(0)
            .WriteDword(0).WriteDword(0).WriteDword(0)
            .WriteNTString("")
            .WriteNTString(text)
            .ToBncsPacket(BncsPacketId.SID_CHATEVENT);

    private static Task InvokeHandleChatEvent(BotEngine engine, byte[] frame)
    {
        var method = typeof(BotEngine).GetMethod("HandleChatEvent", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (Task)method.Invoke(engine, [frame])!;
    }

    [Fact]
    public async Task HandleChatEvent_NoOneHearsYou_LogsRecoveryAttempt()
    {
        var config = new BotConfig { HomeChannel = "Op BNETcc" };
        await using var engine = new BotEngine(config);
        var logged = new List<string>();
        engine.Log += segments => logged.Add(string.Concat(segments.Select(s => s.Text)));

        await InvokeHandleChatEvent(engine, BuildInfoEventFrame("No one hears you."));

        Assert.Contains(logged, l => l.Contains("rejoining home channel", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task HandleChatEvent_NoOneHearsYou_TwiceQuickly_OnlyRecoversOnce()
    {
        var config = new BotConfig { HomeChannel = "Op BNETcc" };
        await using var engine = new BotEngine(config);
        var recoveryLogCount = 0;
        engine.Log += segments =>
        {
            if (string.Concat(segments.Select(s => s.Text)).Contains("rejoining home channel", StringComparison.OrdinalIgnoreCase))
            {
                recoveryLogCount++;
            }
        };

        await InvokeHandleChatEvent(engine, BuildInfoEventFrame("No one hears you."));
        await InvokeHandleChatEvent(engine, BuildInfoEventFrame("No one hears you."));

        Assert.Equal(1, recoveryLogCount);
    }

    [Fact]
    public async Task HandleChatEvent_UnrelatedInfoMessage_DoesNotTriggerRecovery()
    {
        var config = new BotConfig { HomeChannel = "Op BNETcc" };
        await using var engine = new BotEngine(config);
        var logged = new List<string>();
        engine.Log += segments => logged.Add(string.Concat(segments.Select(s => s.Text)));

        await InvokeHandleChatEvent(engine, BuildInfoEventFrame("Welcome to Battle.net."));

        Assert.DoesNotContain(logged, l => l.Contains("rejoining home channel", StringComparison.OrdinalIgnoreCase));
    }
}
