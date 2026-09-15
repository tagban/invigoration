using System.Net;
using Invigoration.Core.Music;
using Invigoration.Core.Music.Spotify;

namespace Invigoration.Core.Tests;

public class SpotifyControllerTests
{
    private const string PlayingTrack = """
        {
          "device": { "id": "d1", "name": "MacBook Pro", "type": "Computer" },
          "progress_ms": 61000,
          "is_playing": true,
          "currently_playing_type": "track",
          "item": {
            "type": "track",
            "name": "Pink Pony Club",
            "uri": "spotify:track:1k2pQc5i348DCHwbn5KTdc",
            "duration_ms": 258000,
            "artists": [ { "name": "Chappell Roan" }, { "name": "Guest" } ],
            "album": {
              "name": "The Rise and Fall of a Midwest Princess",
              "images": [
                { "url": "https://i.scdn.co/image/small", "width": 64, "height": 64 },
                { "url": "https://i.scdn.co/image/large", "width": 640, "height": 640 }
              ]
            }
          }
        }
        """;

    private const string PausedTrack = """{"is_playing":false,"progress_ms":0,"item":{"type":"track","name":"Song","uri":"spotify:track:abc","duration_ms":1000,"artists":[{"name":"Band"}],"album":{"name":"LP","images":[]}}}""";

    private static (SpotifyController Controller, SpotifyTestHttp Http, List<string> Saved) Connect(Func<HttpRequestMessage, HttpResponseMessage> api, SpotifyTokens? tokens = null)
    {
        var saved = new List<string>();
        var http = new SpotifyTestHttp((request, _) => request.RequestUri!.Host == "accounts.spotify.com"
            ? SpotifyTestHttp.Tokens("access-fresh")
            : api(request));
        var controller = new SpotifyController("my-client", "refresh-1", saved.Add, http.Client(),
            tokens ?? new SpotifyTokens("access-1", "refresh-1", DateTimeOffset.UtcNow.AddHours(1)));
        return (controller, http, saved);
    }

    private static List<(HttpMethod Method, string Url)> ApiCalls(SpotifyTestHttp http) =>
        [.. http.Requests.Where(r => r.Url.StartsWith(SpotifyController.ApiBase, StringComparison.Ordinal)).Select(r => (r.Method, r.Url[SpotifyController.ApiBase.Length..]))];

    [Fact]
    public void ParsePlayback_ReadsATrack()
    {
        var info = SpotifyController.ParsePlayback(PlayingTrack)!;

        Assert.Equal("Pink Pony Club", info.Title);
        Assert.Equal("Chappell Roan, Guest", info.Artist);
        Assert.Equal("The Rise and Fall of a Midwest Princess", info.Album);
        Assert.Equal("https://i.scdn.co/image/large", info.ArtworkUrl);
        Assert.True(info.IsPlaying);
        Assert.Equal("MacBook Pro", info.Device);
        Assert.Equal((61000, 258000), (info.ProgressMs, info.DurationMs));
        Assert.Equal("spotify:track:1k2pQc5i348DCHwbn5KTdc", info.Uri);
        Assert.Equal("Spotify", info.Service);
    }

    [Fact]
    public void ParsePlayback_ReadsAPodcastEpisode()
    {
        var info = SpotifyController.ParsePlayback("""{"is_playing":true,"item":{"type":"episode","name":"Episode 12","uri":"spotify:episode:e1","duration_ms":5,"images":[{"url":"https://i/ep","width":300}],"show":{"name":"The Show","images":[]}}}""")!;

        Assert.Equal(("Episode 12", "The Show", "https://i/ep"), (info.Title, info.Artist, info.ArtworkUrl));
    }

    [Theory]
    [InlineData("""{"is_playing":true,"currently_playing_type":"ad","item":null}""")]
    [InlineData("")]
    public void ParsePlayback_NothingLoaded_IsNull(string json)
    {
        Assert.Null(SpotifyController.ParsePlayback(json));
    }

    [Fact]
    public async Task NowPlaying_NoActiveDevice_IsNothingPlaying()
    {
        var (controller, _, _) = Connect(_ => SpotifyTestHttp.Status(HttpStatusCode.NoContent));

        Assert.Null(await controller.GetNowPlayingAsync());
        Assert.Null(controller.LastError);
    }

    [Fact]
    public async Task NowPlaying_IsBrieflyCached_SoABurstOfCommandsAsksSpotifyOnce()
    {
        var (controller, http, _) = Connect(_ => SpotifyTestHttp.Json(PlayingTrack));

        await controller.GetNowPlayingAsync();
        await controller.GetNowPlayingAsync();
        await controller.GetNowPlayingAsync();

        Assert.Single(ApiCalls(http));
        Assert.Equal("Bearer access-1", http.Requests[0].Authorization);
    }

    [Fact]
    public async Task Skip_PostsNext()
    {
        var (controller, http, _) = Connect(_ => SpotifyTestHttp.Status(HttpStatusCode.NoContent));

        Assert.True(await controller.SkipAsync());

        Assert.Equal((HttpMethod.Post, "me/player/next"), Assert.Single(ApiCalls(http)));
    }

    [Fact]
    public async Task PlayPause_PausesWhatsPlaying_AndPlaysWhatsPaused()
    {
        var (playing, playingHttp, _) = Connect(r => r.Method == HttpMethod.Get ? SpotifyTestHttp.Json(PlayingTrack) : SpotifyTestHttp.Status(HttpStatusCode.NoContent));
        Assert.True(await playing.PlayPauseAsync());
        Assert.Equal((HttpMethod.Put, "me/player/pause"), ApiCalls(playingHttp)[^1]);

        var (paused, pausedHttp, _) = Connect(r => r.Method == HttpMethod.Get ? SpotifyTestHttp.Json(PausedTrack) : SpotifyTestHttp.Status(HttpStatusCode.NoContent));
        Assert.True(await paused.PlayPauseAsync());
        Assert.Equal((HttpMethod.Put, "me/player/play"), ApiCalls(pausedHttp)[^1]);
    }

    [Fact]
    public async Task ThumbsUp_SavesTheCurrentTrackThroughTheLibraryEndpoint()
    {
        var (controller, http, _) = Connect(r => r.Method == HttpMethod.Get ? SpotifyTestHttp.Json(PlayingTrack) : SpotifyTestHttp.Status(HttpStatusCode.OK));

        Assert.True(await controller.ThumbsUpAsync());

        Assert.Equal((HttpMethod.Put, "me/library?uris=spotify%3Atrack%3A1k2pQc5i348DCHwbn5KTdc"), ApiCalls(http)[^1]);
    }

    [Fact]
    public async Task ThumbsDown_IsntSomethingSpotifyHas()
    {
        var (controller, http, _) = Connect(_ => throw new InvalidOperationException("no calls expected"));

        Assert.False(controller.SupportsThumbsDown);
        Assert.False(await controller.ThumbsDownAsync());
        Assert.Empty(http.Requests);
    }

    [Theory]
    [InlineData(HttpStatusCode.NotFound, """{"error":{"status":404,"message":"Player command failed: No active device found"}}""", "no active device")]
    [InlineData(HttpStatusCode.Forbidden, """{"error":{"status":403,"message":"Player command failed: Premium required"}}""", "needs Spotify Premium")]
    [InlineData(HttpStatusCode.TooManyRequests, "", "rate-limiting")]
    public async Task AFailedCommand_SaysWhyInWordsChatCanUse(HttpStatusCode status, string body, string expected)
    {
        var (controller, _, _) = Connect(_ => SpotifyTestHttp.Json(body, status));

        Assert.False(await controller.SkipAsync());

        Assert.Contains(expected, controller.LastError);
    }

    [Fact]
    public async Task AnExpiredAccessToken_IsRenewedAndTheCallRetriedOnce()
    {
        var calls = 0;
        var (controller, http, _) = Connect(request => ++calls == 1
            ? SpotifyTestHttp.Status(HttpStatusCode.Unauthorized)
            : SpotifyTestHttp.Status(HttpStatusCode.NoContent));

        Assert.True(await controller.SkipAsync());

        Assert.Equal(2, ApiCalls(http).Count);
        Assert.Equal("Bearer access-fresh", http.Requests[^1].Authorization);
    }

    [Fact]
    public async Task ARenewal_SavesAReplacementRefreshToken()
    {
        var saved = new List<string>();
        var http = new SpotifyTestHttp((request, _) => request.RequestUri!.Host == "accounts.spotify.com"
            ? SpotifyTestHttp.Tokens("access-2", "refresh-2")
            : SpotifyTestHttp.Status(HttpStatusCode.NoContent));
        var controller = new SpotifyController("my-client", "refresh-1", saved.Add, http.Client());

        Assert.True(await controller.SkipAsync());

        Assert.Equal(["refresh-2"], saved);
    }

    [Fact]
    public async Task ARevokedSignIn_IsReportedAsExpired()
    {
        var http = new SpotifyTestHttp((request, _) => request.RequestUri!.Host == "accounts.spotify.com"
            ? SpotifyTestHttp.Json("""{"error":"invalid_grant"}""", HttpStatusCode.BadRequest)
            : throw new InvalidOperationException("the API shouldn't be called without a token"));
        var controller = new SpotifyController("my-client", "refresh-1", _ => { }, http.Client());

        Assert.Null(await controller.GetNowPlayingAsync());

        Assert.True(controller.SignInExpired);
        Assert.Contains("reconnect Spotify", controller.LastError);
    }

    private const string SearchReply = """
        {"tracks":{"items":[
          {"uri":"spotify:track:t1","name":"Pink Pony Club","duration_ms":258000,"artists":[{"name":"Chappell Roan"}],
           "album":{"name":"Midwest Princess","images":[{"url":"https://i/640","width":640},{"url":"https://i/300","width":300},{"url":"https://i/64","width":64}]}},
          null,
          {"uri":"spotify:track:t2","name":"Casual","duration_ms":232000,"artists":[{"name":"Chappell Roan"},{"name":"Other"}],"album":{"name":"Midwest Princess","images":[]}}
        ]}}
        """;

    [Fact]
    public async Task Search_AsksForTracksWithinSpotifysLimit_AndReadsThem()
    {
        var (controller, http, _) = Connect(_ => SpotifyTestHttp.Json(SearchReply));

        var tracks = await controller.SearchTracksAsync("  pink pony club ");

        Assert.Equal((HttpMethod.Get, "search?type=track&limit=10&q=pink%20pony%20club"), Assert.Single(ApiCalls(http)));
        Assert.NotNull(tracks);
        Assert.Equal(2, tracks.Count);
        Assert.Equal(new SpotifyTrack("spotify:track:t1", "Pink Pony Club", "Chappell Roan", "Midwest Princess", "https://i/64", 258000), tracks[0]);
        Assert.Equal(("Chappell Roan, Other", null), (tracks[1].Artist, tracks[1].ArtworkUrl));
    }

    [Fact]
    public async Task Search_ForNothing_DoesntAskSpotify()
    {
        var (controller, http, _) = Connect(_ => throw new InvalidOperationException("no calls expected"));

        Assert.Empty((await controller.SearchTracksAsync("   "))!);
        Assert.Empty(http.Requests);
    }

    [Fact]
    public async Task PlayTrack_SendsTheTracksUri()
    {
        var (controller, http, _) = Connect(_ => SpotifyTestHttp.Status(HttpStatusCode.NoContent));

        Assert.True(await controller.PlayTrackAsync("spotify:track:t1"));

        var request = http.Requests.Single(r => r.Url.StartsWith(SpotifyController.ApiBase, StringComparison.Ordinal));
        Assert.Equal((HttpMethod.Put, SpotifyController.ApiBase + "me/player/play"), (request.Method, request.Url));
        Assert.Equal("""{"uris":["spotify:track:t1"]}""", request.Body);
    }

    [Fact]
    public async Task PlayTrack_WithNothingPlayingAnywhere_PlaysOnAnOpenSpotifyApp()
    {
        const string devices = """{"devices":[{"id":"phone","type":"Smartphone","is_active":false},{"id":"mac","type":"Computer","is_active":false},{"id":"locked","type":"Computer","is_restricted":true}]}""";
        var (controller, http, _) = Connect(r => r.RequestUri!.AbsolutePath switch
        {
            "/v1/me/player/devices" => SpotifyTestHttp.Json(devices),
            _ when r.RequestUri.Query.Contains("device_id") => SpotifyTestHttp.Status(HttpStatusCode.NoContent),
            _ => SpotifyTestHttp.Json("""{"error":{"status":404,"message":"Player command failed: No active device found"}}""", HttpStatusCode.NotFound),
        });

        Assert.True(await controller.PlayTrackAsync("spotify:track:t1"));

        Assert.Equal((HttpMethod.Put, "me/player/play?device_id=mac"), ApiCalls(http)[^1]);
        Assert.Equal("""{"uris":["spotify:track:t1"]}""", http.Requests[^1].Body);
        Assert.False(controller.NoDeviceOpen);
    }

    [Fact]
    public async Task PlayTrack_WithNoSpotifyAppOpenAtAll_SaysSo()
    {
        var (controller, _, _) = Connect(r => r.RequestUri!.AbsolutePath == "/v1/me/player/devices"
            ? SpotifyTestHttp.Json("""{"devices":[]}""")
            : SpotifyTestHttp.Json("""{"error":{"status":404,"message":"No active device found"}}""", HttpStatusCode.NotFound));

        Assert.False(await controller.PlayTrackAsync("spotify:track:t1"));

        Assert.True(controller.NoDeviceOpen);
        Assert.False(await controller.HasOpenDeviceAsync());
        Assert.Contains("No Spotify app is open", controller.LastError);
    }

    [Fact]
    public async Task QueueTrack_AddsItNext_OrPlaysItWhenThereIsNoQueue()
    {
        var (queued, queuedHttp, _) = Connect(_ => SpotifyTestHttp.Status(HttpStatusCode.NoContent));
        Assert.True(await queued.QueueTrackAsync("spotify:track:t1"));
        Assert.Equal((HttpMethod.Post, "me/player/queue?uri=spotify%3Atrack%3At1"), Assert.Single(ApiCalls(queuedHttp)));

        var (idle, idleHttp, _) = Connect(r => r.RequestUri!.AbsolutePath.EndsWith("/queue", StringComparison.Ordinal)
            ? SpotifyTestHttp.Status(HttpStatusCode.NotFound)
            : SpotifyTestHttp.Status(HttpStatusCode.NoContent));
        Assert.True(await idle.QueueTrackAsync("spotify:track:t1"));
        Assert.Equal((HttpMethod.Put, "me/player/play"), ApiCalls(idleHttp)[^1]);
    }

    [Fact]
    public async Task PlayPause_WithNothingLoaded_ResumesTheLastSession()
    {
        var (controller, http, _) = Connect(r => r.Method == HttpMethod.Get
            ? SpotifyTestHttp.Status(HttpStatusCode.NoContent)
            : SpotifyTestHttp.Status(HttpStatusCode.NoContent));

        Assert.True(await controller.PlayPauseAsync());

        Assert.Equal((HttpMethod.Put, "me/player/play"), ApiCalls(http)[^1]);
        Assert.Equal("", http.Requests[^1].Body);
    }
}

[Collection(MusicPlayerRegistryCollection.Name)]
public class MusicNowPlayingReplyTests
{
    private sealed class Fake(NowPlayingInfo? nowPlaying, string? lastError = null) : IMusicPlayerController
    {
        public Task<bool> SkipAsync() => Task.FromResult(false);
        public Task<bool> PlayPauseAsync() => Task.FromResult(false);
        public Task<bool> ThumbsUpAsync() => Task.FromResult(false);
        public Task<bool> ThumbsDownAsync() => Task.FromResult(false);
        public Task<NowPlayingInfo?> GetNowPlayingAsync() => Task.FromResult(nowPlaying);
        public string? LastError => lastError;
    }

    private static async Task<string> Describe(IMusicPlayerController? controller)
    {
        MusicPlayerRegistry.Controller = controller;
        try
        {
            return await MusicPlayerRegistry.DescribeNowPlayingAsync();
        }
        finally
        {
            MusicPlayerRegistry.Controller = null;
        }
    }

    [Fact]
    public async Task Playing_Paused_Nothing_AndNotConnected()
    {
        Assert.Equal("/me is now playing Song - by Band on Spotify.", await Describe(new Fake(new NowPlayingInfo("Song", "Band", "Spotify"))));
        Assert.Equal("/me has Song - by Band paused on Spotify.", await Describe(new Fake(new NowPlayingInfo("Song", "Band", "Spotify") { IsPlaying = false })));
        Assert.Equal("Nothing seems to be playing.", await Describe(new Fake(null)));
        Assert.Equal("Couldn't reach Spotify.", await Describe(new Fake(null, "Couldn't reach Spotify.")));
        Assert.Equal(MusicPlayerRegistry.NotConnectedReply, await Describe(null));
    }
}
