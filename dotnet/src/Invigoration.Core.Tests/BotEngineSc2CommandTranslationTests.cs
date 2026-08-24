using System.Reflection;
using Invigoration.Core.Config;
using Invigoration.Core.Protocol;
using Stimpak;

namespace Invigoration.Core.Tests;

public class Sc2EmoteTranslationTests
{
    [Fact]
    public void MeText_IsWrappedInAsterisksInsteadOfSentLiterally()
    {
        Assert.Equal("*is an Invigoration v2.0.2b - bnet.cc*", BotEngine.TranslateSc2EmoteText("/me is an Invigoration v2.0.2b - bnet.cc"));
    }

    [Fact]
    public void OrdinaryText_PassesThroughUnchanged()
    {
        Assert.Equal("hello there", BotEngine.TranslateSc2EmoteText("hello there"));
    }

    [Fact]
    public void SlashMeWithNoTrailingSpace_PassesThroughUnchanged()
    {
        // "/mean" etc. shouldn't be mistaken for the "/me " emote prefix.
        Assert.Equal("/mean", BotEngine.TranslateSc2EmoteText("/mean"));
    }
}

/// <summary>
/// Regression coverage for a real bug: SC2 whisper replies used to send the literal text
/// "/w username message" as a public channel message instead of an actual whisper, since
/// nothing intercepted the universal "/w " convention for Stimpak-backed products the way
/// SendSc2Async's TryParseSc2Whisper now does.
/// </summary>
public class Sc2WhisperParsingTests
{
    [Fact]
    public void StandardWhisperBody_ParsesTargetAndMessage()
    {
        Assert.True(BotEngine.TryParseSc2Whisper("/w SomePlayer hey there", out var target, out var message));
        Assert.Equal("SomePlayer", target);
        Assert.Equal("hey there", message);
    }

    [Fact]
    public void MessageWithSpaces_KeepsThemIntact()
    {
        Assert.True(BotEngine.TryParseSc2Whisper("/w SomePlayer this has several words in it", out _, out var message));
        Assert.Equal("this has several words in it", message);
    }

    [Fact]
    public void OrdinaryChatText_IsNotParsedAsAWhisper()
    {
        Assert.False(BotEngine.TryParseSc2Whisper("hello everyone", out _, out _));
    }

    [Fact]
    public void MissingMessageBody_IsNotParsedAsAWhisper()
    {
        // "/w SomePlayer" with nothing after it — no space to split target from message.
        Assert.False(BotEngine.TryParseSc2Whisper("/w SomePlayer", out _, out _));
    }
}

/// <summary>
/// Covers Config.Sc2LastChannels staying in sync with actually-joined channels — see
/// BotEngine.Sc2.cs's PersistSc2ChannelList. Channel *restoration* on connect is now handled
/// natively by Stimpak itself (StimpakConnectOptions.Channels, passed in ConnectSc2Async), not
/// something this engine replays by hand any more, so there's nothing left to test for that
/// half — only that we keep an accurate list for the next connect to hand back to Stimpak.
/// </summary>
public class BotEngineSc2ChannelPersistenceTests
{
    private static BotEngine NewSc2Engine(out BotConfig config)
    {
        // See BotEngineSc2ChannelStateTests.NewSc2Engine's matching comment — bypassing
        // ConnectSc2Async means this registration (normally done there) has to happen here too.
        StimpakNativeResolver.Register();

        config = new BotConfig { Product = BncsProduct.Sc2, DisplayName = $"bot-{Guid.NewGuid():N}" };
        var engine = new BotEngine(config);

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

    [Fact]
    public async Task JoiningAPublicChannel_AddsItToConfigAndFiresConfigPersistNeeded()
    {
        await using var engine = NewSc2Engine(out var config);
        var fired = 0;
        engine.ConfigPersistNeeded += () => fired++;

        await InvokeHandleSc2Event(engine, new Joined(1, new PublicChannel(100, "General"), 1));

        Assert.Equal([ChannelTarget.Public(100)], config.Sc2LastChannels);
        Assert.Equal(1, fired);
    }

    [Fact]
    public async Task JoiningAPrivateChannel_AddsItByName()
    {
        await using var engine = NewSc2Engine(out var config);

        await InvokeHandleSc2Event(engine, new Joined(1, new PrivateChannel("Clan BNU"), 1));

        Assert.Equal([ChannelTarget.Private("Clan BNU")], config.Sc2LastChannels);
    }

    [Fact]
    public async Task LeavingAChannel_RemovesItFromConfig()
    {
        await using var engine = NewSc2Engine(out var config);

        await InvokeHandleSc2Event(engine, new Joined(1, new PublicChannel(100, "General"), 1));
        await InvokeHandleSc2Event(engine, new Joined(2, new PrivateChannel("Clan BNU"), 1));
        await InvokeHandleSc2Event(engine, new Left(1, null));

        Assert.Equal([ChannelTarget.Private("Clan BNU")], config.Sc2LastChannels);
    }
}
