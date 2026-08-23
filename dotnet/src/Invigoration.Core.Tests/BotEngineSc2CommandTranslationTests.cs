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

/// <summary>Covers Config.Sc2LastChannelNames staying in sync with actually-joined channels, so a later reconnect can restore the same set — see BotEngine.Sc2.cs's RejoinRememberedSc2Channels/PersistSc2ChannelList.</summary>
public class BotEngineSc2ChannelPersistenceTests
{
    private static BotEngine NewSc2Engine(out BotConfig config)
    {
        config = new BotConfig { Product = BncsProduct.Sc2, DisplayName = $"bot-{Guid.NewGuid():N}" };
        var engine = new BotEngine(config);

        var credentialPath = Path.Combine(Path.GetTempPath(), $"stimpak-test-{Guid.NewGuid():N}.bin");
        var client = new StimpakClient(credentialPath);
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
    public async Task JoiningAChannel_AddsItsNameToConfigAndFiresConfigPersistNeeded()
    {
        await using var engine = NewSc2Engine(out var config);
        var fired = 0;
        engine.ConfigPersistNeeded += () => fired++;

        await InvokeHandleSc2Event(engine, new Joined(1, new PublicChannel(100, "General"), 1));

        Assert.Equal(["General"], config.Sc2LastChannelNames);
        Assert.Equal(1, fired);
    }

    [Fact]
    public async Task LeavingAChannel_RemovesItsNameFromConfig()
    {
        await using var engine = NewSc2Engine(out var config);

        await InvokeHandleSc2Event(engine, new Joined(1, new PublicChannel(100, "General"), 1));
        await InvokeHandleSc2Event(engine, new Joined(2, new PrivateChannel("Clan BNU"), 1));
        await InvokeHandleSc2Event(engine, new Left(1, null));

        Assert.Equal(["Clan BNU"], config.Sc2LastChannelNames);
    }

    [Fact]
    public async Task RejoinRememberedChannels_SkipsANameThatIsAlreadyJoined()
    {
        // The default channel auto-joins unconditionally on connect; if it happens to match a
        // remembered name, replay must not attempt a redundant (and native-call-triggering)
        // second join for it.
        await using var engine = NewSc2Engine(out var config);
        config.Sc2LastChannelNames = ["General"];
        await InvokeHandleSc2Event(engine, new Joined(1, new PublicChannel(100, "General"), 1));

        var fired = 0;
        engine.ConfigPersistNeeded += () => fired++;

        var method = typeof(BotEngine).GetMethod("RejoinRememberedSc2Channels", BindingFlags.NonPublic | BindingFlags.Instance)!;
        method.Invoke(engine, []);

        // Only "General" is remembered and it's already joined, so nothing should have changed.
        Assert.Equal(["General"], config.Sc2LastChannelNames);
        Assert.Equal(0, fired);
    }

    /// <summary>
    /// Regression test for a real bug: replay used to run as soon as PublicChannelsReceived
    /// arrived, with no guarantee the always-auto-joined default channel's own Joined
    /// confirmation had landed first. If it hadn't, replay couldn't see the default channel in
    /// _sc2Channels yet and re-attempted joining it — the server correctly rejected the
    /// duplicate, surfacing a confusing "Could not join General" error even though the bot was
    /// already in General. MaybeRejoinRememberedSc2Channels now gates on both preconditions and
    /// only ever runs once per connection.
    /// </summary>
    [Fact]
    public async Task MaybeRejoinRememberedChannels_OnlyRunsOnceBothPreconditionsAreMetAndNeverTwice()
    {
        await using var engine = NewSc2Engine(out var config);
        config.Sc2LastChannelNames = ["General"];
        var attemptedField = typeof(BotEngine).GetField("_sc2RejoinAttempted", BindingFlags.NonPublic | BindingFlags.Instance)!;

        // Default channel confirmed, but the catalog hasn't arrived yet — must not attempt yet.
        await InvokeHandleSc2Event(engine, new Joined(1, new PublicChannel(100, "General"), 1));
        Assert.False((bool)attemptedField.GetValue(engine)!);

        // Catalog arrives — both preconditions are now met, so the guard flips (the actual
        // replay attempt for "General" is a no-op since it's already joined).
        await InvokeHandleSc2Event(engine, new PublicChannelsReceived([new PublicChannel(100, "General")]));
        Assert.True((bool)attemptedField.GetValue(engine)!);

        // A second PublicChannelsReceived must not re-run replay.
        var namesBefore = config.Sc2LastChannelNames.ToList();
        await InvokeHandleSc2Event(engine, new PublicChannelsReceived([new PublicChannel(100, "General")]));
        Assert.Equal(namesBefore, config.Sc2LastChannelNames);
    }
}
