using System.Reflection;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>
/// Covers the "never triggerable from chat, only locally" restriction added 2026-08-26 for
/// config-mutating commands and cosmetic text-transform toggles (BotEngine.Commands.cs's
/// LocalOnlyCommands) — checked before IsAuthorized, so even the bot master typing one of these
/// from chat (not this app's own local input box) is silently ignored.
/// </summary>
public class BotEngineLocalOnlyCommandTests
{
    private static Task InvokeRemoteCommand(BotEngine engine, string username, string message)
    {
        var method = typeof(BotEngine).GetMethod("HandleCommandAsync", BindingFlags.NonPublic | BindingFlags.Instance,
            null, [typeof(string), typeof(string), typeof(bool), typeof(byte?)], null)!;
        return (Task)method.Invoke(engine, [username, message, false, null])!;
    }

    [Fact]
    public async Task SetPass_FromChat_EvenAsBotMaster_IsIgnored()
    {
        var config = new BotConfig { Trigger = "!", BotMaster = "TheMaster", Password = "original" };
        await using var engine = new BotEngine(config);

        await InvokeRemoteCommand(engine, "TheMaster", "!setpass hijacked");

        Assert.Equal("original", config.Password);
    }

    [Fact]
    public async Task SetPass_RunLocally_StillWorks()
    {
        var config = new BotConfig { Trigger = "!", BotMaster = "TheMaster", Password = "original" };
        await using var engine = new BotEngine(config);

        await engine.RunLocalCommandAsync("/setpass changed");

        Assert.Equal("changed", config.Password);
    }

    [Theory]
    [InlineData("colors")]
    [InlineData("canada")]
    [InlineData("debug")]
    [InlineData("settrigger")]
    public async Task LocalOnlyCommand_FromChat_DoesNotThrow(string command)
    {
        var config = new BotConfig { Trigger = "!", BotMaster = "TheMaster" };
        await using var engine = new BotEngine(config);

        // Just confirming these are silently swallowed (no exception, no crash) rather than
        // asserting on every individual side effect — SetPass above already covers the actual
        // "state doesn't change" guarantee in detail.
        await InvokeRemoteCommand(engine, "TheMaster", $"!{command} arg");
    }
}
