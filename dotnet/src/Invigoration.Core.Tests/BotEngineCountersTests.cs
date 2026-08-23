using System.Reflection;
using Invigoration.Core.Chat;
using Invigoration.Core.Commands;
using Invigoration.Core.Config;
using Invigoration.Core.Protocol;

namespace Invigoration.Core.Tests;

/// <summary>
/// Covers a real bug found via live testing: BanCount/KickCount/JoinCount
/// (the "!bancount"/"!kickcount"/"!joincount" commands) were never actually
/// incremented anywhere in the port — only read. Fixed by wiring
/// BotEngine.Bncs.cs's HandleChatEvent to increment them the same way the
/// VB6 original did (ChatBot_OnJoin/OnInfo/OnChannel in frmMain.frm):
/// JoinCount on every ChatEventType.Join, BanCount/KickCount on
/// ChatEventType.Info messages containing "was banned by"/"was kicked out of
/// the channel by", and all three reset to 0 whenever the bot (re)joins a
/// channel (ChatEventType.Channel).
/// </summary>
public class BotEngineCountersTests
{
    private static byte[] BuildFrame(ChatEventType type, string username, string text) =>
        new PacketWriter()
            .WriteDword((uint)type)
            .WriteDword(0)
            .WriteDword(0)
            .WriteDword(0).WriteDword(0).WriteDword(0)
            .WriteNTString(username)
            .WriteNTString(text)
            .ToBncsPacket(BncsPacketId.SID_CHATEVENT);

    private static Task InvokeHandleChatEvent(BotEngine engine, byte[] frame)
    {
        var method = typeof(BotEngine).GetMethod("HandleBncsChatEventFrame", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (Task)method.Invoke(engine, [frame])!;
    }

    private static BotSessionState GetSession(BotEngine engine)
    {
        var field = typeof(BotEngine).GetField("_session", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (BotSessionState)field.GetValue(engine)!;
    }

    [Fact]
    public async Task HandleChatEvent_Join_IncrementsJoinCount()
    {
        var config = new BotConfig();
        await using var engine = new BotEngine(config);

        await InvokeHandleChatEvent(engine, BuildFrame(ChatEventType.Join, "SomeUser", "PX2D0000"));
        await InvokeHandleChatEvent(engine, BuildFrame(ChatEventType.Join, "OtherUser", "PX2D0000"));

        Assert.Equal(2, GetSession(engine).JoinCount);
    }

    [Fact]
    public async Task HandleChatEvent_ShowUser_DoesNotIncrementJoinCount()
    {
        var config = new BotConfig();
        await using var engine = new BotEngine(config);

        await InvokeHandleChatEvent(engine, BuildFrame(ChatEventType.ShowUser, "SomeUser", "PX2D0000"));

        Assert.Equal(0, GetSession(engine).JoinCount);
    }

    [Fact]
    public async Task HandleChatEvent_KickedInfoMessage_IncrementsKickCount()
    {
        var config = new BotConfig();
        await using var engine = new BotEngine(config);

        await InvokeHandleChatEvent(engine, BuildFrame(ChatEventType.Info, "", "Someone was kicked out of the channel by SomeOp."));

        Assert.Equal(1, GetSession(engine).KickCount);
        Assert.Equal(0, GetSession(engine).BanCount);
    }

    [Fact]
    public async Task HandleChatEvent_BannedInfoMessage_IncrementsBanCount()
    {
        var config = new BotConfig();
        await using var engine = new BotEngine(config);

        await InvokeHandleChatEvent(engine, BuildFrame(ChatEventType.Info, "", "Someone was banned by SomeOp."));

        Assert.Equal(1, GetSession(engine).BanCount);
        Assert.Equal(0, GetSession(engine).KickCount);
    }

    [Fact]
    public async Task HandleChatEvent_JoiningNewChannel_ResetsAllCounts()
    {
        var config = new BotConfig();
        await using var engine = new BotEngine(config);
        var session = GetSession(engine);
        session.JoinCount = 5;
        session.BanCount = 3;
        session.KickCount = 2;

        await InvokeHandleChatEvent(engine, BuildFrame(ChatEventType.Channel, "", "Op BNETcc"));

        Assert.Equal(0, session.JoinCount);
        Assert.Equal(0, session.BanCount);
        Assert.Equal(0, session.KickCount);
    }
}
