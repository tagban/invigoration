using System.Reflection;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;
using Invigoration.Core.Protocol;
using Stimpak;

namespace Invigoration.Core.Tests;

public class BotEngineSc2ChannelStateTests
{
    private static BotEngine NewSc2Engine()
    {
        // These tests construct a StimpakClient directly (below), bypassing
        // BotEngine.ConnectSc2Async entirely — that's normally where this registration
        // happens, so it needs to happen here too, or the very first StimpakClient
        // construction in the whole test run throws (the OS finds Stimpak's own managed
        // assembly before ever probing runtimes/<rid>/native/ — see StimpakNativeResolver).
        StimpakNativeResolver.Register();

        var config = new BotConfig { Product = BncsProduct.Sc2, DisplayName = $"bot-{Guid.NewGuid():N}" };
        var engine = new BotEngine(config);

        // Give it a live StimpakClient (never connected) so HandleSc2EventAsync's
        // unconditional client.People.Apply(next) has something real to call —
        // this is the same native library the app itself links against.
        var credentialPath = Path.Combine(Path.GetTempPath(), $"stimpak-test-{Guid.NewGuid():N}.bin");
        var client = new StimpakClient(new StimpakClientOptions("cc.bnet.invigoration.tests") { CredentialPath = credentialPath });
        typeof(BotEngine).GetField("_sc2Client", BindingFlags.NonPublic | BindingFlags.Instance)!.SetValue(engine, client);

        return engine;
    }

    private static Task InvokeHandleSc2Event(BotEngine engine, SC2Event next)
    {
        var client = (StimpakClient)typeof(BotEngine).GetField("_sc2Client", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(engine)!;
        var method = typeof(BotEngine).GetMethod("HandleSc2EventAsync", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (Task)method.Invoke(engine, [client, next])!;
    }

    private static Dictionary<byte, object> GetChannels(BotEngine engine)
    {
        var field = typeof(BotEngine).GetField("_sc2Channels", BindingFlags.NonPublic | BindingFlags.Instance)!;
        var dict = field.GetValue(engine)!;
        var keys = (System.Collections.IEnumerable)dict.GetType().GetProperty("Keys")!.GetValue(dict)!;
        var result = new Dictionary<byte, object>();
        var indexer = dict.GetType().GetProperty("Item")!;
        foreach (byte key in keys)
        {
            result[key] = indexer.GetValue(dict, [key])!;
        }

        return result;
    }

    private static void SetChannelCount(BotEngine engine, int count)
    {
        var field = typeof(BotEngine).GetField("_sc2Channels", BindingFlags.NonPublic | BindingFlags.Instance)!;
        var dict = field.GetValue(engine)!;
        var addMethod = dict.GetType().GetMethod("Add")!;
        var sessionType = typeof(BotEngine).GetNestedType("Sc2ChannelSession", BindingFlags.NonPublic)!;
        var publicChannelCtor = typeof(PublicChannel).GetConstructor([typeof(ushort), typeof(string)])!;
        for (var i = 0; i < count; i++)
        {
            var channel = publicChannelCtor.Invoke([(ushort)i, $"Channel{i}"]);
            var session = Activator.CreateInstance(sessionType, [(byte)i, channel]);
            addMethod.Invoke(dict, [(byte)i, session]);
        }
    }

    [Fact]
    public async Task MemberJoined_DifferentChannels_DoNotAffectEachOther()
    {
        await using var engine = NewSc2Engine();
        var events = new List<ChatEvent>();
        engine.ChatMessage += events.Add;

        await InvokeHandleSc2Event(engine, new Joined(1, new PublicChannel(100, "A"), 1));
        await InvokeHandleSc2Event(engine, new Joined(2, new PublicChannel(200, "B"), 1));
        await InvokeHandleSc2Event(engine, new MemberJoined(2, new User(5, null, "Someone", null, Presence.Available)));

        var joinEvents = events.Where(e => e.Type == ChatEventType.Join).ToList();
        Assert.Single(joinEvents);
        Assert.Equal((byte)2, joinEvents[0].ChannelIndex);
    }

    [Fact]
    public async Task MessageReceived_CarriesTheOriginatingChannelIndex()
    {
        await using var engine = NewSc2Engine();
        var events = new List<ChatEvent>();
        engine.ChatMessage += events.Add;

        await InvokeHandleSc2Event(engine, new Joined(3, new PublicChannel(300, "C"), 1));
        await InvokeHandleSc2Event(engine, new MessageReceived(3, new User(5, null, "Someone", null, Presence.Available), "hello"));

        var talk = Assert.Single(events, e => e.Type == ChatEventType.Talk);
        Assert.Equal((byte)3, talk.ChannelIndex);
        Assert.Equal("hello", talk.Text);
    }

    [Fact]
    public async Task Joined_OnlyFiresBncsConnectedOnceAcrossMultipleChannels()
    {
        await using var engine = NewSc2Engine();
        var connectedCount = 0;
        engine.BncsConnected += () => connectedCount++;

        await InvokeHandleSc2Event(engine, new Joined(1, new PublicChannel(100, "A"), 1));
        await InvokeHandleSc2Event(engine, new Joined(2, new PublicChannel(200, "B"), 1));

        Assert.Equal(1, connectedCount);
        Assert.Equal(2, GetChannels(engine).Count);
    }

    [Fact]
    public async Task Left_RemovesOnlyThatChannel()
    {
        await using var engine = NewSc2Engine();
        var left = new List<byte>();
        engine.Sc2ChannelLeft += left.Add;

        await InvokeHandleSc2Event(engine, new Joined(1, new PublicChannel(100, "A"), 1));
        await InvokeHandleSc2Event(engine, new Joined(2, new PublicChannel(200, "B"), 1));
        await InvokeHandleSc2Event(engine, new Left(1, null));

        Assert.Equal([(byte)1], left);
        Assert.DoesNotContain((byte)1, GetChannels(engine).Keys);
        Assert.Contains((byte)2, GetChannels(engine).Keys);
    }

    /// <summary>
    /// Regression test for a real bug: LeaveSc2Channel used to only send the native leave call
    /// and then wait for an async Left SC2Event to actually remove the channel/close the tab.
    /// Stimpak's own leave_channel (native/superiority/core/src/games/sc2/chat/session.rs)
    /// forgets the channel from its local state synchronously as part of that same call — the
    /// Left event is driven by a separate, server-pushed roster update that isn't guaranteed to
    /// follow a leave the client itself initiated, which left some tabs (most visibly the
    /// always-auto-joined default channel) stuck open indefinitely. LeaveSc2Channel now removes
    /// its own tracking immediately after a successful native call, with no Left event needed.
    /// </summary>
    [Fact]
    public async Task LeaveSc2Channel_RemovesTheChannelImmediately_WithNoLeftEventNeeded()
    {
        await using var engine = NewSc2Engine();
        var left = new List<byte>();
        engine.Sc2ChannelLeft += left.Add;

        await InvokeHandleSc2Event(engine, new Joined(1, new PublicChannel(100, "General"), 1));
        engine.LeaveSc2Channel(1);

        Assert.Equal([(byte)1], left);
        Assert.DoesNotContain((byte)1, GetChannels(engine).Keys);
    }

    [Fact]
    public async Task SessionEnded_FiresLeftForEveryTrackedChannelAndClears()
    {
        await using var engine = NewSc2Engine();
        var left = new List<byte>();
        engine.Sc2ChannelLeft += left.Add;

        await InvokeHandleSc2Event(engine, new Joined(1, new PublicChannel(100, "A"), 1));
        await InvokeHandleSc2Event(engine, new Joined(2, new PublicChannel(200, "B"), 1));
        await InvokeHandleSc2Event(engine, new SessionEnded());

        Assert.Equal([(byte)1, (byte)2], left.OrderBy(b => b).ToArray());
        Assert.Empty(GetChannels(engine));
    }

    /// <summary>
    /// Regression test for a real bug: WhisperReceived's Outgoing:true case (Stimpak's
    /// synchronous local echo confirming a whisper this bot itself sent — see
    /// send_resolved_whisper in the Rust core) was never handled at all, only Outgoing:false
    /// (incoming). A whisper reply on SC2 was sent correctly but never showed up anywhere in
    /// this bot's own UI, since nothing ever produced a WhisperSent ChatEvent for it.
    /// </summary>
    [Fact]
    public async Task WhisperReceived_Outgoing_ProducesAWhisperSentChatEvent()
    {
        await using var engine = NewSc2Engine();
        var events = new List<ChatEvent>();
        engine.ChatMessage += events.Add;

        await InvokeHandleSc2Event(engine, new WhisperReceived("SomePlayer", "hey there", true));

        var sent = Assert.Single(events);
        Assert.Equal(ChatEventType.WhisperSent, sent.Type);
        Assert.Equal("SomePlayer", sent.Username);
        Assert.Equal("hey there", sent.Text);
    }

    [Fact]
    public async Task WhisperReceived_NotOutgoing_ProducesAWhisperChatEvent()
    {
        await using var engine = NewSc2Engine();
        var events = new List<ChatEvent>();
        engine.ChatMessage += events.Add;

        await InvokeHandleSc2Event(engine, new WhisperReceived("SomePlayer", "hi", false));

        var received = Assert.Single(events);
        Assert.Equal(ChatEventType.Whisper, received.Type);
        Assert.Equal("SomePlayer", received.Username);
    }

    [Fact]
    public async Task TryJoinSc2PublicChannel_AtCap_ReturnsFalseAndFiresRejectionWithNoLiveClient()
    {
        var config = new BotConfig { Product = BncsProduct.Sc2 };
        await using var engine = new BotEngine(config);
        SetChannelCount(engine, BotEngine.MaxJoinedSc2Channels);
        string? rejection = null;
        engine.Sc2ChannelJoinRejected += reason => rejection = reason;

        var result = engine.TryJoinSc2PublicChannel(999);

        Assert.False(result);
        Assert.NotNull(rejection);
        Assert.Contains(BotEngine.MaxJoinedSc2Channels.ToString(), rejection);
    }

    [Fact]
    public async Task TryJoinSc2PublicChannel_NonSc2Product_ReturnsFalse()
    {
        var config = new BotConfig { Product = BncsProduct.Starcraft };
        await using var engine = new BotEngine(config);

        Assert.False(engine.TryJoinSc2PublicChannel(1028));
    }
}
