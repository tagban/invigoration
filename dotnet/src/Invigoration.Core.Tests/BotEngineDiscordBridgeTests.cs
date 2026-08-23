using System.Reflection;
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
