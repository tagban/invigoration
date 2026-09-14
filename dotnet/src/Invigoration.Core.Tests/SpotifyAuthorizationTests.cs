using System.Net;
using System.Net.Sockets;
using Invigoration.Core.Music.Spotify;

namespace Invigoration.Core.Tests;

public class SpotifyAuthorizationTests
{
    // RFC 7636, Appendix B.
    [Fact]
    public void CodeChallenge_MatchesThePkceSpecExample()
    {
        Assert.Equal("E9Melhoa2OwvFrEMTJguCHaoeK1t8URWbuGJSstw-cM",
            SpotifyAuthorization.CodeChallenge("dBjftJeZ4CVP-mB92K27uhbUJU1p1r_wW1gFWFOEjXk"));
    }

    [Fact]
    public void CodeVerifier_IsUrlSafeAndLongEnough()
    {
        var verifier = SpotifyAuthorization.CreateCodeVerifier();

        Assert.InRange(verifier.Length, 43, 128);
        Assert.Matches("^[A-Za-z0-9_-]+$", verifier);
        Assert.NotEqual(verifier, SpotifyAuthorization.CreateCodeVerifier());
    }

    [Fact]
    public void AuthorizeUri_CarriesEverySpotifyAsksFor()
    {
        var uri = SpotifyAuthorization.AuthorizeUri("my-client", "the-challenge", "the-state");
        var query = System.Web.HttpUtility.ParseQueryString(uri.Query);

        Assert.Equal("https://accounts.spotify.com/authorize", uri.GetLeftPart(UriPartial.Path));
        Assert.Equal("code", query["response_type"]);
        Assert.Equal("my-client", query["client_id"]);
        Assert.Equal("http://127.0.0.1:43117/callback", query["redirect_uri"]);
        Assert.Equal("S256", query["code_challenge_method"]);
        Assert.Equal("the-challenge", query["code_challenge"]);
        Assert.Equal("the-state", query["state"]);
        Assert.Equal("user-read-playback-state user-read-currently-playing user-modify-playback-state user-library-modify", query["scope"]);
    }

    [Fact]
    public async Task ExchangeCode_PostsThePkceFormAndReadsTheTokens()
    {
        var http = new SpotifyTestHttp((_, _) => SpotifyTestHttp.Tokens("access-1", "refresh-1", 3600));

        var tokens = await SpotifyAuthorization.ExchangeCodeAsync(http.Client(), "my-client", "the-code", "the-verifier");

        var (method, url, body, _) = Assert.Single(http.Requests);
        Assert.Equal(HttpMethod.Post, method);
        Assert.Equal("https://accounts.spotify.com/api/token", url);
        var form = System.Web.HttpUtility.ParseQueryString(body);
        Assert.Equal("authorization_code", form["grant_type"]);
        Assert.Equal("the-code", form["code"]);
        Assert.Equal("http://127.0.0.1:43117/callback", form["redirect_uri"]);
        Assert.Equal("my-client", form["client_id"]);
        Assert.Equal("the-verifier", form["code_verifier"]);
        Assert.Equal(("access-1", "refresh-1"), (tokens.AccessToken, tokens.RefreshToken));
        Assert.InRange(tokens.ExpiresAt - DateTimeOffset.UtcNow, TimeSpan.FromMinutes(59), TimeSpan.FromMinutes(61));
    }

    [Fact]
    public async Task Refresh_KeepsTheOldRefreshTokenWhenSpotifyDoesntSendANewOne()
    {
        var http = new SpotifyTestHttp((_, _) => SpotifyTestHttp.Tokens("access-2"));

        var tokens = await SpotifyAuthorization.RefreshAsync(http.Client(), "my-client", "refresh-1");

        var form = System.Web.HttpUtility.ParseQueryString(Assert.Single(http.Requests).Body);
        Assert.Equal("refresh_token", form["grant_type"]);
        Assert.Equal("refresh-1", form["refresh_token"]);
        Assert.Equal("my-client", form["client_id"]);
        Assert.Equal(("access-2", "refresh-1"), (tokens.AccessToken, tokens.RefreshToken));
    }

    [Fact]
    public async Task Refresh_ARevokedSignIn_SaysOnlyReconnectingHelps()
    {
        var http = new SpotifyTestHttp((_, _) => SpotifyTestHttp.Json("""{"error":"invalid_grant","error_description":"Refresh token revoked"}""", HttpStatusCode.BadRequest));

        var ex = await Assert.ThrowsAsync<SpotifyAuthException>(() => SpotifyAuthorization.RefreshAsync(http.Client(), "my-client", "refresh-1"));

        Assert.True(ex.SignInExpired);
        Assert.Contains("Refresh token revoked", ex.Message);
    }

    private static int FreePort()
    {
        using var probe = new TcpListener(IPAddress.Loopback, 0);
        probe.Start();
        return ((IPEndPoint)probe.LocalEndpoint).Port;
    }

    [Fact(Timeout = 10000)]
    public async Task WaitForCode_ReturnsTheCodeFromTheBrowsersReply()
    {
        var port = FreePort();
        var waiting = SpotifyAuthorization.WaitForCodeAsync("state-1", port);

        using var browser = new HttpClient();
        Assert.Equal(HttpStatusCode.NotFound, (await browser.GetAsync($"http://127.0.0.1:{port}/favicon.ico")).StatusCode);
        var page = await browser.GetStringAsync($"http://127.0.0.1:{port}/callback?code=the-code&state=state-1");

        Assert.Equal("the-code", await waiting);
        Assert.Contains("connected to Spotify", page);
    }

    [Theory(Timeout = 10000)]
    [InlineData("code=the-code&state=someone-elses", "for this request")] // the page HTML-encodes the apostrophe in "wasn't"
    [InlineData("error=access_denied&state=state-1", "cancelled")]
    public async Task WaitForCode_RefusesAReplyThatIsntASuccessfulSignIn(string query, string expected)
    {
        var port = FreePort();
        var waiting = SpotifyAuthorization.WaitForCodeAsync("state-1", port);

        using var browser = new HttpClient();
        var page = await browser.GetStringAsync($"http://127.0.0.1:{port}/callback?{query}");

        var ex = await Assert.ThrowsAsync<SpotifyAuthException>(() => waiting);
        Assert.Contains(expected, ex.Message);
        Assert.Contains(expected, page);
    }

    [Fact(Timeout = 10000)]
    public async Task WaitForCode_StopsWhenCancelled()
    {
        using var cts = new CancellationTokenSource();
        var waiting = SpotifyAuthorization.WaitForCodeAsync("state-1", FreePort(), cts.Token);

        cts.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => waiting);
    }
}
