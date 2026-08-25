using System.Reflection;
using Invigoration.Core.Config;
using Invigoration.Core.Music;

namespace Invigoration.Core.Tests;

/// <summary>
/// Covers the "!skip"/"!thumbsup"/"!thumbsdown"/"!nowplaying" chat commands
/// (BotEngine.Commands.cs) against a fake IMusicPlayerController — the real one
/// (YouTubeMusicWindow, Invigoration.App) drives an actual embedded browser and isn't
/// unit-testable, but the dispatch/null-handling logic here is. MusicPlayerRegistry.Controller
/// is process-wide static state; every test resets it in a finally block so a failure here can't
/// leak into other test files (none currently touch it, but cheap insurance regardless — see
/// the BattlenetCredentialProfileStoreTests fixture for why shared static state left dirty is a
/// real, previously-hit problem in this codebase).
/// </summary>
public class BotEngineMusicCommandTests
{
    private sealed class FakeMusicPlayerController : IMusicPlayerController
    {
        public string? LastCalled { get; private set; }
        public bool NextResult { get; set; } = true;
        public NowPlayingInfo? NowPlaying { get; set; }

        public Task<bool> SkipAsync()
        {
            LastCalled = "skip";
            return Task.FromResult(NextResult);
        }

        public Task<bool> ThumbsUpAsync()
        {
            LastCalled = "thumbsup";
            return Task.FromResult(NextResult);
        }

        public Task<bool> ThumbsDownAsync()
        {
            LastCalled = "thumbsdown";
            return Task.FromResult(NextResult);
        }

        public Task<NowPlayingInfo?> GetNowPlayingAsync()
        {
            LastCalled = "nowplaying";
            return Task.FromResult(NowPlaying);
        }

        public bool SupportsThumbsDown { get; set; } = true;
    }

    private static Task InvokeRemoteCommand(BotEngine engine, string username, string message)
    {
        var method = typeof(BotEngine).GetMethod("HandleCommandAsync", BindingFlags.NonPublic | BindingFlags.Instance,
            null, [typeof(string), typeof(string), typeof(bool), typeof(byte?)], null)!;
        return (Task)method.Invoke(engine, [username, message, false, null])!;
    }

    private static BotEngine CreateBotMasterEngine() => new(new BotConfig { Trigger = "!", BotMaster = "TheMaster" });

    [Fact]
    public async Task Skip_WithControllerRegistered_CallsSkipAsync()
    {
        var fake = new FakeMusicPlayerController();
        MusicPlayerRegistry.Controller = fake;
        try
        {
            await using var engine = CreateBotMasterEngine();
            await InvokeRemoteCommand(engine, "TheMaster", "!skip");
            Assert.Equal("skip", fake.LastCalled);
        }
        finally
        {
            MusicPlayerRegistry.Controller = null;
        }
    }

    [Fact]
    public async Task Next_IsAnAliasForSkip()
    {
        var fake = new FakeMusicPlayerController();
        MusicPlayerRegistry.Controller = fake;
        try
        {
            await using var engine = CreateBotMasterEngine();
            await InvokeRemoteCommand(engine, "TheMaster", "!next");
            Assert.Equal("skip", fake.LastCalled);
        }
        finally
        {
            MusicPlayerRegistry.Controller = null;
        }
    }

    [Fact]
    public async Task ThumbsUp_WithControllerRegistered_CallsThumbsUpAsync()
    {
        var fake = new FakeMusicPlayerController();
        MusicPlayerRegistry.Controller = fake;
        try
        {
            await using var engine = CreateBotMasterEngine();
            await InvokeRemoteCommand(engine, "TheMaster", "!thumbsup");
            Assert.Equal("thumbsup", fake.LastCalled);
        }
        finally
        {
            MusicPlayerRegistry.Controller = null;
        }
    }

    [Fact]
    public async Task ThumbsDown_WithControllerRegistered_CallsThumbsDownAsync()
    {
        var fake = new FakeMusicPlayerController();
        MusicPlayerRegistry.Controller = fake;
        try
        {
            await using var engine = CreateBotMasterEngine();
            await InvokeRemoteCommand(engine, "TheMaster", "!thumbsdown");
            Assert.Equal("thumbsdown", fake.LastCalled);
        }
        finally
        {
            MusicPlayerRegistry.Controller = null;
        }
    }

    [Fact]
    public async Task ThumbsDown_WhenServiceDoesNotSupportIt_QuietlyDoesNothing()
    {
        var fake = new FakeMusicPlayerController { SupportsThumbsDown = false };
        MusicPlayerRegistry.Controller = fake;
        try
        {
            await using var engine = CreateBotMasterEngine();
            await InvokeRemoteCommand(engine, "TheMaster", "!thumbsdown");
            Assert.Null(fake.LastCalled);
        }
        finally
        {
            MusicPlayerRegistry.Controller = null;
        }
    }

    [Fact]
    public async Task NowPlaying_WithControllerRegistered_CallsGetNowPlayingAsync()
    {
        var fake = new FakeMusicPlayerController { NowPlaying = new NowPlayingInfo("Song", "Artist") };
        MusicPlayerRegistry.Controller = fake;
        try
        {
            await using var engine = CreateBotMasterEngine();
            await InvokeRemoteCommand(engine, "TheMaster", "!nowplaying");
            Assert.Equal("nowplaying", fake.LastCalled);
        }
        finally
        {
            MusicPlayerRegistry.Controller = null;
        }
    }

    [Fact]
    public async Task Np_IsAnAliasForNowPlaying()
    {
        var fake = new FakeMusicPlayerController { NowPlaying = new NowPlayingInfo("Song", "Artist") };
        MusicPlayerRegistry.Controller = fake;
        try
        {
            await using var engine = CreateBotMasterEngine();
            await InvokeRemoteCommand(engine, "TheMaster", "!np");
            Assert.Equal("nowplaying", fake.LastCalled);
        }
        finally
        {
            MusicPlayerRegistry.Controller = null;
        }
    }

    [Fact]
    public async Task Skip_WithNoControllerRegistered_DoesNotThrow()
    {
        MusicPlayerRegistry.Controller = null;
        await using var engine = CreateBotMasterEngine();

        // "Music player isn't open." reply just goes out over a disconnected, no-op wire send in
        // tests — the point here is that a null controller doesn't throw, not the reply text.
        await InvokeRemoteCommand(engine, "TheMaster", "!skip");
    }

    [Fact]
    public async Task Skip_WhenActionReportsFailure_DoesNotThrow()
    {
        var fake = new FakeMusicPlayerController { NextResult = false };
        MusicPlayerRegistry.Controller = fake;
        try
        {
            await using var engine = CreateBotMasterEngine();
            await InvokeRemoteCommand(engine, "TheMaster", "!skip");
            Assert.Equal("skip", fake.LastCalled);
        }
        finally
        {
            MusicPlayerRegistry.Controller = null;
        }
    }
}
