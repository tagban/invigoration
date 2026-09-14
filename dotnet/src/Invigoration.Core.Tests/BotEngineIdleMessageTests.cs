using System.Reflection;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>Covers ResolveIdlePlaceholdersAsync (BotEngine.Idle.cs) — the actual idle-timer trigger itself isn't unit-tested here (it's a 30s-interval background loop), but the placeholder substitution it depends on is.</summary>
[Collection(MusicPlayerRegistryCollection.Name)]
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

    private sealed class NowPlayingOnly(Music.NowPlayingInfo nowPlaying) : Music.IMusicPlayerController
    {
        public Task<bool> SkipAsync() => Task.FromResult(false);
        public Task<bool> PlayPauseAsync() => Task.FromResult(false);
        public Task<bool> ThumbsUpAsync() => Task.FromResult(false);
        public Task<bool> ThumbsDownAsync() => Task.FromResult(false);
        public Task<Music.NowPlayingInfo?> GetNowPlayingAsync() => Task.FromResult<Music.NowPlayingInfo?>(nowPlaying);
    }

    [Theory]
    [InlineData(true, "Now playing: Song by Band!")]
    [InlineData(false, "Now playing: !")] // a paused track isn't "playing"
    public async Task ResolveIdlePlaceholders_MusicPlaying_OnlyWhileItsPlaying(bool isPlaying, string expected)
    {
        Music.MusicPlayerRegistry.Controller = new NowPlayingOnly(new Music.NowPlayingInfo("Song", "Band", "Spotify") { IsPlaying = isPlaying });
        try
        {
            await using var engine = new BotEngine(new BotConfig());
            Assert.Equal(expected, await ResolveAsync(engine, "Now playing: %MusicPlaying%!"));
        }
        finally
        {
            Music.MusicPlayerRegistry.Controller = null;
        }
    }
}
