using System.Reflection;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>Covers ResolveIdlePlaceholdersAsync (BotEngine.Idle.cs) — the actual idle-timer trigger itself isn't unit-tested here (it's a 30s-interval background loop), but the placeholder substitution it depends on is.</summary>
public class BotEngineIdleMessageTests
{
    private static Task<string> ResolveAsync(BotEngine engine, string template)
    {
        var method = typeof(BotEngine).GetMethod("ResolveIdlePlaceholdersAsync", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (Task<string>)method.Invoke(engine, [template])!;
    }

    [Fact]
    public async Task ResolveIdlePlaceholders_SubstitutesVerAndUsername()
    {
        var config = new BotConfig { Username = "TestBot" };
        await using var engine = new BotEngine(config);

        var resolved = await ResolveAsync(engine, "I am %Username%, version %Ver%.");

        Assert.Contains("I am TestBot, version", resolved);
        Assert.DoesNotContain("%Username%", resolved);
        Assert.DoesNotContain("%Ver%", resolved);
    }

    [Fact]
    public async Task ResolveIdlePlaceholders_IsCaseInsensitive()
    {
        var config = new BotConfig { Username = "TestBot" };
        await using var engine = new BotEngine(config);

        var resolved = await ResolveAsync(engine, "%USERNAME% / %username% / %UserName%");

        Assert.Equal("TestBot / TestBot / TestBot", resolved);
    }

    [Fact]
    public async Task ResolveIdlePlaceholders_MusicPlayingWithNoControllerRegistered_BecomesEmpty()
    {
        // MusicPlayerRegistry.Controller is process-wide static state — explicitly null it so this
        // test doesn't depend on whatever another test file left behind (see
        // BotEngineMusicCommandTests' remarks on the same shared-state concern).
        Music.MusicPlayerRegistry.Controller = null;
        var config = new BotConfig();
        await using var engine = new BotEngine(config);

        var resolved = await ResolveAsync(engine, "Now playing: %MusicPlaying%!");

        Assert.Equal("Now playing: !", resolved);
    }

    [Fact]
    public async Task ResolveIdlePlaceholders_LeavesTemplateWithNoPlaceholdersUnchanged()
    {
        var config = new BotConfig();
        await using var engine = new BotEngine(config);

        var resolved = await ResolveAsync(engine, "back in a bit");

        Assert.Equal("back in a bit", resolved);
    }
}
