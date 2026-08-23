using System.Reflection;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;
using Invigoration.Core.Protocol;

namespace Invigoration.Core.Tests;

/// <summary>
/// Covers a user-requested feature, demonstrated live by a real server
/// (eurobattle.net) where one user's flaky client bounced in and out of the
/// channel dozens of times in under a second — HideJoinLeaveSpamEnabled lets
/// the chat log stop showing further Join/Leave lines for that one user once
/// they cross the configured rate, without touching anything else (roster
/// tracking, JoinCount, rank behaviors all still run underneath).
/// </summary>
[Collection("ClanRosterStore")]
public class BotEngineJoinLeaveSpamTests
{
    private static byte[] BuildFrame(ChatEventType type, string username) =>
        new PacketWriter()
            .WriteDword((uint)type)
            .WriteDword(0)
            .WriteDword(0)
            .WriteDword(0).WriteDword(0).WriteDword(0)
            .WriteNTString(username)
            .WriteNTString("PX2D")
            .ToBncsPacket(BncsPacketId.SID_CHATEVENT);

    private static Task InvokeHandleChatEvent(BotEngine engine, byte[] frame)
    {
        var method = typeof(BotEngine).GetMethod("HandleBncsChatEventFrame", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (Task)method.Invoke(engine, [frame])!;
    }

    private static int GetSessionJoinCount(BotEngine engine)
    {
        var field = typeof(BotEngine).GetField("_session", BindingFlags.NonPublic | BindingFlags.Instance)!;
        var session = field.GetValue(engine)!;
        return (int)session.GetType().GetProperty("JoinCount")!.GetValue(session)!;
    }

    [Fact]
    public async Task HideJoinLeaveSpamEnabled_UserExceedsThreshold_FurtherJoinLeaveLinesAreHiddenFromChatLog()
    {
        var config = new BotConfig { HideJoinLeaveSpamEnabled = true, HideJoinLeaveSpamThreshold = 3, HideJoinLeaveSpamWindowSeconds = 60 };
        await using var engine = new BotEngine(config);
        var shownEvents = new List<ChatEvent>();
        engine.ChatMessage += e => shownEvents.Add(e);

        // 8 join/leave events in a row for the same flaky user — matches the real capture.
        for (var i = 0; i < 4; i++)
        {
            await InvokeHandleChatEvent(engine, BuildFrame(ChatEventType.Join, "shadowmoon"));
            await InvokeHandleChatEvent(engine, BuildFrame(ChatEventType.Leave, "shadowmoon"));
        }

        // First 3 join + 3 leave (interleaved) should still show; everything past the
        // threshold should be silently dropped from the visible log.
        Assert.True(shownEvents.Count < 8, $"Expected some events to be suppressed, but all {shownEvents.Count} showed.");
        Assert.All(shownEvents, e => Assert.True(e.Type is ChatEventType.Join or ChatEventType.Leave));

        // The underlying join count must NOT be affected by the display filter — every
        // Join still increments it, hidden or not.
        Assert.Equal(4, GetSessionJoinCount(engine));
    }

    [Fact]
    public async Task HideJoinLeaveSpamEnabled_DifferentUsers_EachTrackedIndependently()
    {
        var config = new BotConfig { HideJoinLeaveSpamEnabled = true, HideJoinLeaveSpamThreshold = 1, HideJoinLeaveSpamWindowSeconds = 60 };
        await using var engine = new BotEngine(config);
        var shownEvents = new List<ChatEvent>();
        engine.ChatMessage += e => shownEvents.Add(e);

        await InvokeHandleChatEvent(engine, BuildFrame(ChatEventType.Join, "alice"));
        await InvokeHandleChatEvent(engine, BuildFrame(ChatEventType.Join, "bob"));

        // Threshold is 1 ("more than 1"), so each user's very first join alone should still show.
        Assert.Equal(2, shownEvents.Count);
    }

    [Fact]
    public async Task HideJoinLeaveSpamDisabled_NeverSuppressesAnything()
    {
        var config = new BotConfig { HideJoinLeaveSpamEnabled = false, HideJoinLeaveSpamThreshold = 1 };
        await using var engine = new BotEngine(config);
        var shownEvents = new List<ChatEvent>();
        engine.ChatMessage += e => shownEvents.Add(e);

        for (var i = 0; i < 10; i++)
        {
            await InvokeHandleChatEvent(engine, BuildFrame(ChatEventType.Join, "shadowmoon"));
        }

        Assert.Equal(10, shownEvents.Count);
    }

    [Fact]
    public async Task NonJoinLeaveEvents_AreNeverSuppressedRegardlessOfRate()
    {
        var config = new BotConfig { HideJoinLeaveSpamEnabled = true, HideJoinLeaveSpamThreshold = 1 };
        await using var engine = new BotEngine(config);
        var shownEvents = new List<ChatEvent>();
        engine.ChatMessage += e => shownEvents.Add(e);

        for (var i = 0; i < 5; i++)
        {
            await InvokeHandleChatEvent(engine, BuildFrame(ChatEventType.Talk, "shadowmoon"));
        }

        Assert.Equal(5, shownEvents.Count);
    }
}
