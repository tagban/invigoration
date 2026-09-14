using System.Reflection;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>
/// The Discord bridge itself (gateway connection, message relay) needs a
/// real bot token and Discord server to test against, neither of which is
/// available in this environment — so this only covers what's testable
/// without one: that a disabled/unconfigured bridge never even attempts to
/// start, so every other BotEngine test that connects (with Discord left at
/// its default-off config) isn't secretly trying to reach Discord's gateway.
/// </summary>
public class BotEngineDiscordBridgeTests
{
    private static void InvokeStartIfEnabled(BotEngine engine)
    {
        var method = typeof(BotEngine).GetMethod("StartDiscordBridgeIfEnabled", BindingFlags.NonPublic | BindingFlags.Instance)!;
        method.Invoke(engine, []);
    }

    private static object? GetBridge(BotEngine engine) =>
        typeof(BotEngine).GetField("_discordBridge", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(engine);

    [Fact]
    public async Task StartDiscordBridgeIfEnabled_DisabledByDefault_NeverStarts()
    {
        var config = new BotConfig();
        await using var engine = new BotEngine(config);

        InvokeStartIfEnabled(engine);

        Assert.Null(GetBridge(engine));
    }

    [Fact]
    public async Task StartDiscordBridgeIfEnabled_EnabledButNoToken_NeverStarts()
    {
        var config = new BotConfig();
        config.Discord.Enabled = true;
        config.Discord.BotToken = "";
        await using var engine = new BotEngine(config);

        InvokeStartIfEnabled(engine);

        Assert.Null(GetBridge(engine));
    }
}

/// <summary>
/// A message arriving from Discord: shown once as "[Discord] name", relayed to Battle.net without a
/// second local copy, never posted back to Discord, and never able to borrow a Battle.net identity.
/// Driven through the private handler by reflection (no live Discord connection), on an engine
/// marked logged on and in a channel so its sends go through the normal path.
/// </summary>
public class BotEngineDiscordRelayTests
{
    private static Task Deliver(BotEngine engine, string discordUser, string content) =>
        (Task)typeof(BotEngine).GetMethod("HandleDiscordMessageAsync", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .Invoke(engine, [discordUser, content])!;

    private static bool ShouldRelayToDiscord(ChatEvent e, bool enabled) =>
        (bool)typeof(BotEngine).GetMethod("ShouldRelayToDiscord", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!
            .Invoke(null, [e, enabled])!;

    private static BotConfig RelayConfig() => new()
    {
        Username = "BNU",
        BotMaster = "Tagban",
        Trigger = "!",
        FloodProtectionDelayMs = 0,
        Discord = new DiscordBridgeConfig { Enabled = true, RelayDelaySeconds = 0 },
    };

    [Fact]
    public async Task ADiscordMessage_IsShownOnceAsDiscordName_WithNoSecondCopyFromTheRelay()
    {
        await using var engine = BotEngineChatGateTests.MarkInChannel(new BotEngine(RelayConfig()));
        var shown = new List<ChatEvent>();
        var echoed = new List<string>();
        engine.ChatMessage += shown.Add;
        engine.SelfChatSent += s => echoed.Add(string.Concat(s.Select(x => x.Text)));

        await Deliver(engine, "dave", "hello from discord");

        var line = Assert.Single(shown);
        Assert.Equal("[Discord] dave", line.Username);
        Assert.Equal("hello from discord", line.Text);
        Assert.Equal(ChatEventOrigin.Discord, line.Origin);
        Assert.Empty(echoed);
    }

    [Fact]
    public void MessagesFromDiscord_AreNeverRelayedBackToDiscord()
    {
        Assert.False(ShouldRelayToDiscord(new ChatEvent(ChatEventType.Talk, "[Discord] dave", 0, 0, "hi", Origin: ChatEventOrigin.Discord), enabled: true));
        Assert.True(ShouldRelayToDiscord(new ChatEvent(ChatEventType.Talk, "Tagban", 0, 0, "hi"), enabled: true));
        Assert.True(ShouldRelayToDiscord(new ChatEvent(ChatEventType.Emote, "Tagban", 0, 0, "waves"), enabled: true));
        Assert.False(ShouldRelayToDiscord(new ChatEvent(ChatEventType.Join, "Tagban", 0, 0, ""), enabled: true));
        Assert.False(ShouldRelayToDiscord(new ChatEvent(ChatEventType.Talk, "Tagban", 0, 0, "hi"), enabled: false));
    }

    // Regression: Discord names are self-chosen, so a Discord account named after the bot master
    // used to be treated as the bot master and could run protected commands through the relay.
    [Fact]
    public async Task ADiscordUserNamedLikeTheBotMaster_CantRunProtectedCommands()
    {
        await using var engine = BotEngineChatGateTests.MarkInChannel(new BotEngine(RelayConfig()));
        var echoed = new List<string>();
        engine.SelfChatSent += s => echoed.Add(string.Concat(s.Select(x => x.Text)));

        await Deliver(engine, "Tagban", "!say pwned");

        Assert.DoesNotContain(echoed, e => e.Contains("pwned", StringComparison.Ordinal) && !e.Contains("[Discord]", StringComparison.Ordinal));
        Assert.Empty(echoed);
    }

    // The same command from the real bot master on Battle.net still works — the fix is about who's speaking, not the command.
    [Fact]
    public async Task TheRealBotMaster_OnBattlenet_StillCan()
    {
        await using var engine = BotEngineChatGateTests.MarkInChannel(new BotEngine(RelayConfig()));
        var echoed = new List<string>();
        engine.SelfChatSent += s => echoed.Add(string.Concat(s.Select(x => x.Text)));
        var handle = typeof(BotEngine).GetMethod("HandleChatEvent", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!;

        await (Task)handle.Invoke(engine, [new ChatEvent(ChatEventType.Talk, "Tagban", 0, 0, "!say hi everyone")])!;

        Assert.Contains(echoed, e => e.EndsWith("hi everyone", StringComparison.Ordinal));
    }
}
