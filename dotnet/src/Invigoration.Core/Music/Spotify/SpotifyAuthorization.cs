using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace Invigoration.Core.Music.Spotify;

/// <summary>A Spotify sign-in or token refresh that didn't work. <see cref="SignInExpired"/> means the saved sign-in was revoked or has expired, so only connecting again will help.</summary>
public sealed class SpotifyAuthException(string message, bool signInExpired = false) : Exception(message)
{
    public bool SignInExpired { get; } = signInExpired;
}

/// <summary>An access token and the refresh token that renews it.</summary>
public sealed record SpotifyTokens(string AccessToken, string RefreshToken, DateTimeOffset ExpiresAt);

/// <summary>
/// Spotify sign-in: the Authorization Code flow with PKCE, which a desktop app can use without a
/// client secret. Each user registers their own Spotify developer app and pastes its Client ID —
/// development-mode apps are limited to a handful of users (five, since February 2026), so one
/// shared Invigoration app couldn't serve everyone. The browser signs in and comes back to a
/// loopback address on this machine (<see cref="RedirectUri"/>): Spotify only allows plain HTTP
/// for an explicit loopback IP, never "localhost".
/// </summary>
public static class SpotifyAuthorization
{
    public const int CallbackPort = 43117;

    /// <summary>What the user registers as their Spotify app's redirect URI, exactly.</summary>
    public const string RedirectUri = "http://127.0.0.1:43117/callback";

    public const string AuthorizeEndpoint = "https://accounts.spotify.com/authorize";
    public const string TokenEndpoint = "https://accounts.spotify.com/api/token";

    /// <summary>Read what's playing, control playback, and save tracks (the like button).</summary>
    public static IReadOnlyList<string> Scopes { get; } =
        ["user-read-playback-state", "user-read-currently-playing", "user-modify-playback-state", "user-library-modify"];

    /// <summary>A PKCE code verifier: 64 URL-safe characters from 48 random bytes.</summary>
    public static string CreateCodeVerifier() => Base64Url(RandomNumberGenerator.GetBytes(48));

    /// <summary>The S256 challenge for a verifier: base64url(SHA-256(verifier)).</summary>
    public static string CodeChallenge(string codeVerifier) => Base64Url(SHA256.HashData(Encoding.ASCII.GetBytes(codeVerifier)));

    public static string CreateState() => Base64Url(RandomNumberGenerator.GetBytes(16));

    public static Uri AuthorizeUri(string clientId, string codeChallenge, string state, string redirectUri = RedirectUri)
    {
        var query = string.Join("&", new[]
        {
            ("response_type", "code"),
            ("client_id", clientId),
            ("redirect_uri", redirectUri),
            ("code_challenge_method", "S256"),
            ("code_challenge", codeChallenge),
            ("scope", string.Join(' ', Scopes)),
            ("state", state),
        }.Select(p => $"{p.Item1}={Uri.EscapeDataString(p.Item2)}"));
        return new Uri($"{AuthorizeEndpoint}?{query}");
    }

    /// <summary>Trades the code the browser brought back for tokens.</summary>
    public static Task<SpotifyTokens> ExchangeCodeAsync(HttpClient http, string clientId, string code, string codeVerifier, string redirectUri = RedirectUri, CancellationToken cancellationToken = default) =>
        RequestTokensAsync(http, new Dictionary<string, string>
        {
            ["grant_type"] = "authorization_code",
            ["code"] = code,
            ["redirect_uri"] = redirectUri,
            ["client_id"] = clientId,
            ["code_verifier"] = codeVerifier,
        }, previousRefreshToken: null, cancellationToken);

    /// <summary>Gets a fresh access token. Spotify only sometimes sends a new refresh token; when it doesn't, the old one stays in use.</summary>
    public static Task<SpotifyTokens> RefreshAsync(HttpClient http, string clientId, string refreshToken, CancellationToken cancellationToken = default) =>
        RequestTokensAsync(http, new Dictionary<string, string>
        {
            ["grant_type"] = "refresh_token",
            ["refresh_token"] = refreshToken,
            ["client_id"] = clientId,
        }, refreshToken, cancellationToken);

    /// <summary>
    /// Waits for the browser to come back from Spotify to <see cref="RedirectUri"/> and returns the
    /// authorization code, after checking it answers this sign-in (<paramref name="expectedState"/>).
    /// The browser tab gets a short page saying whether it worked.
    /// </summary>
    public static async Task<string> WaitForCodeAsync(string expectedState, int port = CallbackPort, CancellationToken cancellationToken = default)
    {
        using var listener = new HttpListener();
        listener.Prefixes.Add($"http://127.0.0.1:{port}/");
        try
        {
            listener.Start();
        }
        catch (HttpListenerException ex)
        {
            throw new SpotifyAuthException($"Couldn't listen for Spotify's sign-in reply on port {port}: {ex.Message}");
        }

        using var stop = cancellationToken.Register(listener.Stop);
        while (true)
        {
            HttpListenerContext context;
            try
            {
                context = await listener.GetContextAsync().ConfigureAwait(false);
            }
            catch (Exception ex) when ((ex is HttpListenerException or ObjectDisposedException) && cancellationToken.IsCancellationRequested)
            {
                throw new OperationCanceledException(cancellationToken);
            }

            if (context.Request.Url?.AbsolutePath != "/callback")
            {
                context.Response.StatusCode = 404;
                context.Response.Close();
                continue;
            }

            var query = context.Request.QueryString;
            string? failure = null;
            if (query["state"] != expectedState)
            {
                failure = "That sign-in reply wasn't for this request. Try connecting again.";
            }
            else if (query["error"] is { } error)
            {
                failure = error == "access_denied" ? "Spotify sign-in was cancelled." : $"Spotify sign-in failed: {error}.";
            }
            else if (string.IsNullOrEmpty(query["code"]))
            {
                failure = "Spotify didn't send a sign-in code.";
            }

            await RespondAsync(context, failure is null
                ? "Invigoration is connected to Spotify. You can close this tab."
                : failure).ConfigureAwait(false);

            return failure is null ? query["code"]! : throw new SpotifyAuthException(failure);
        }
    }

    private static async Task RespondAsync(HttpListenerContext context, string message)
    {
        var html = $"<!doctype html><meta charset=\"utf-8\"><title>Invigoration</title>" +
                   $"<body style=\"font-family:system-ui;background:#121212;color:#fff;display:grid;place-items:center;height:90vh\">" +
                   $"<p style=\"font-size:18px\">{WebUtility.HtmlEncode(message)}</p></body>";
        var bytes = Encoding.UTF8.GetBytes(html);
        context.Response.ContentType = "text/html; charset=utf-8";
        context.Response.ContentLength64 = bytes.Length;
        await context.Response.OutputStream.WriteAsync(bytes).ConfigureAwait(false);
        context.Response.Close();
    }

    private static async Task<SpotifyTokens> RequestTokensAsync(HttpClient http, Dictionary<string, string> form, string? previousRefreshToken, CancellationToken cancellationToken)
    {
        HttpResponseMessage response;
        try
        {
            response = await http.PostAsync(TokenEndpoint, new FormUrlEncodedContent(form), cancellationToken).ConfigureAwait(false);
        }
        catch (HttpRequestException ex)
        {
            throw new SpotifyAuthException($"Couldn't reach Spotify: {ex.Message}");
        }

        using (response)
        {
            var body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            JsonElement root;
            try
            {
                root = JsonDocument.Parse(body).RootElement;
            }
            catch (JsonException)
            {
                throw new SpotifyAuthException($"Spotify sent an unreadable reply ({(int)response.StatusCode}).");
            }

            if (!response.IsSuccessStatusCode)
            {
                var error = root.TryGetProperty("error", out var e) ? e.GetString() : null;
                var description = root.TryGetProperty("error_description", out var d) ? d.GetString() : null;
                // invalid_grant on a refresh: the user removed access, or the token expired. Only a new sign-in fixes that.
                throw new SpotifyAuthException(
                    $"Spotify refused the sign-in: {description ?? error ?? response.StatusCode.ToString()}.",
                    signInExpired: error == "invalid_grant" && previousRefreshToken is not null);
            }

            var accessToken = root.GetProperty("access_token").GetString() ?? "";
            var refreshToken = root.TryGetProperty("refresh_token", out var r) && r.GetString() is { Length: > 0 } fresh
                ? fresh
                : previousRefreshToken ?? throw new SpotifyAuthException("Spotify didn't send a refresh token.");
            var expiresIn = root.TryGetProperty("expires_in", out var x) && x.TryGetInt32(out var seconds) ? seconds : 3600;
            return new SpotifyTokens(accessToken, refreshToken, DateTimeOffset.UtcNow.AddSeconds(expiresIn));
        }
    }

    private static string Base64Url(byte[] bytes) => Convert.ToBase64String(bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_');
}
