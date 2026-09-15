using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace Invigoration.Core.Music.Spotify;

/// <summary>
/// The bot's music commands, carried out on the user's own Spotify through the Spotify Web API:
/// Spotify plays wherever the user already listens (the desktop app, a phone, a speaker) and
/// Invigoration reads and controls it. Nothing plays inside Invigoration itself, so there's no
/// embedded web player to break. Playback control needs Spotify Premium; reading what's playing
/// doesn't.
/// </summary>
/// <remarks>
/// Shared by every bot at once, so it's safe to call from several threads: the access token is
/// renewed under a lock, and what's playing is cached for a moment so a title bar, a player bar
/// and a burst of !np commands don't each ask Spotify separately. <see cref="LastError"/> says why
/// the most recent action didn't work, in words a chat reply can use.
/// </remarks>
public sealed class SpotifyController : IMusicPlayerController
{
    public const string ServiceName = "Spotify";
    public const string ApiBase = "https://api.spotify.com/v1/";

    /// <summary>The most results one search can return: Spotify's cap for development-mode apps since February 2026.</summary>
    public const int MaxSearchResults = 10;

    private static readonly TimeSpan NowPlayingCacheLifetime = TimeSpan.FromSeconds(2);

    private readonly HttpClient _http;
    private readonly string _clientId;
    private readonly Action<string> _saveRefreshToken;
    private readonly SemaphoreSlim _tokenLock = new(1, 1);
    private SpotifyTokens? _tokens;
    private string _refreshToken;
    private (DateTimeOffset At, NowPlayingInfo? Info)? _nowPlayingCache;

    /// <param name="clientId">The user's own Spotify app Client ID.</param>
    /// <param name="refreshToken">The saved sign-in.</param>
    /// <param name="saveRefreshToken">Called whenever Spotify hands out a replacement refresh token, so the saved sign-in stays valid.</param>
    /// <param name="http">For tests; one shared client otherwise.</param>
    /// <param name="tokens">Tokens already in hand (straight after signing in), so the first call doesn't renew them.</param>
    public SpotifyController(string clientId, string refreshToken, Action<string> saveRefreshToken, HttpClient? http = null, SpotifyTokens? tokens = null)
    {
        _clientId = clientId;
        _refreshToken = tokens?.RefreshToken ?? refreshToken;
        _saveRefreshToken = saveRefreshToken;
        _http = http ?? SharedHttp;
        _tokens = tokens;
    }

    /// <summary>One HttpClient for every Spotify call Invigoration makes.</summary>
    public static HttpClient SharedHttp { get; } = new() { Timeout = TimeSpan.FromSeconds(15) };

    /// <summary>Why the last action didn't work, or null if it did.</summary>
    public string? LastError { get; private set; }

    /// <summary>True once Spotify has revoked or expired the saved sign-in: nothing will work again until the user reconnects.</summary>
    public bool SignInExpired { get; private set; }

    /// <summary>True after an action failed because none of the user's Spotify apps was open to play on — the cue to open one (see the Music tab's Start).</summary>
    public bool NoDeviceOpen { get; private set; }

    /// <summary>Spotify has "save to your library" (the heart) but no dislike.</summary>
    public bool SupportsThumbsDown => false;

    public async Task<NowPlayingInfo?> GetNowPlayingAsync()
    {
        if (_nowPlayingCache is { } cached && DateTimeOffset.UtcNow - cached.At < NowPlayingCacheLifetime)
        {
            return cached.Info;
        }

        var (info, ok) = await FetchNowPlayingAsync().ConfigureAwait(false);
        if (ok)
        {
            _nowPlayingCache = (DateTimeOffset.UtcNow, info);
        }

        return info;
    }

    public Task<bool> SkipAsync() => SendAsync(HttpMethod.Post, "me/player/next");

    public async Task<bool> PlayPauseAsync()
    {
        var (current, ok) = await FetchNowPlayingAsync().ConfigureAwait(false);
        if (!ok)
        {
            return false;
        }

        // Nothing loaded anywhere: resume the last session on whichever Spotify app is open.
        return current is null
            ? await ResumeAsync().ConfigureAwait(false)
            : await SendAsync(HttpMethod.Put, current.IsPlaying ? "me/player/pause" : "me/player/play").ConfigureAwait(false);
    }

    /// <summary>Saves the current track (or episode) to the user's Spotify library — Spotify's heart.</summary>
    public async Task<bool> ThumbsUpAsync()
    {
        var (current, ok) = await FetchNowPlayingAsync().ConfigureAwait(false);
        if (!ok)
        {
            return false;
        }

        if (current?.Uri is not { Length: > 0 } uri)
        {
            LastError = "Nothing is playing on Spotify to save.";
            return false;
        }

        // February 2026: the per-type save endpoints (/me/tracks and friends) were replaced by this one, which takes URIs.
        return await SendAsync(HttpMethod.Put, $"me/library?uris={Uri.EscapeDataString(uri)}").ConfigureAwait(false);
    }

    public Task<bool> ThumbsDownAsync()
    {
        LastError = "Spotify doesn't have a dislike.";
        return Task.FromResult(false);
    }

    /// <summary>Finds tracks by name, artist or both. Null (with LastError set) when Spotify couldn't be asked.</summary>
    public async Task<IReadOnlyList<SpotifyTrack>?> SearchTracksAsync(string query)
    {
        query = query.Trim();
        if (query.Length == 0)
        {
            return [];
        }

        using var response = await CallAsync(HttpMethod.Get, $"search?type=track&limit={MaxSearchResults}&q={Uri.EscapeDataString(query)}").ConfigureAwait(false);
        if (response is null)
        {
            return null;
        }

        if (!response.IsSuccessStatusCode)
        {
            LastError = await DescribeFailureAsync(response).ConfigureAwait(false);
            return null;
        }

        try
        {
            var tracks = ParseTrackSearch(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
            LastError = null;
            return tracks;
        }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException or KeyNotFoundException)
        {
            LastError = "Spotify sent something unexpected.";
            return null;
        }
    }

    /// <summary>Resumes playback — the last session, on the active device or else an open Spotify app.</summary>
    public Task<bool> ResumeAsync() => SendToDeviceAsync(HttpMethod.Put, "me/player/play", null);

    /// <summary>Whether any of the user's Spotify apps is open and can be told to play.</summary>
    public async Task<bool> HasOpenDeviceAsync() => await FindDeviceAsync().ConfigureAwait(false) is not null;

    /// <summary>Plays a track now, in place of whatever's playing. With nothing playing anywhere, starts it on one of the user's open Spotify apps.</summary>
    public Task<bool> PlayTrackAsync(string uri) =>
        SendToDeviceAsync(HttpMethod.Put, "me/player/play", JsonSerializer.Serialize(new { uris = new[] { uri } }));

    /// <summary>Adds a track to play next. With nothing playing anywhere there's no queue to add to, so it just plays.</summary>
    public async Task<bool> QueueTrackAsync(string uri)
    {
        using var response = await CallAsync(HttpMethod.Post, $"me/player/queue?uri={Uri.EscapeDataString(uri)}").ConfigureAwait(false);
        if (response is null)
        {
            return false;
        }

        _nowPlayingCache = null;
        if (response.IsSuccessStatusCode)
        {
            LastError = null;
            return true;
        }

        if (response.StatusCode == HttpStatusCode.NotFound)
        {
            return await PlayTrackAsync(uri).ConfigureAwait(false);
        }

        LastError = await DescribeFailureAsync(response).ConfigureAwait(false);
        return false;
    }

    /// <summary>Reads a GET /search?type=track reply.</summary>
    public static IReadOnlyList<SpotifyTrack> ParseTrackSearch(string json)
    {
        using var document = JsonDocument.Parse(json);
        if (!document.RootElement.TryGetProperty("tracks", out var tracks) ||
            !tracks.TryGetProperty("items", out var items) || items.ValueKind != JsonValueKind.Array)
        {
            return [];
        }

        var results = new List<SpotifyTrack>();
        foreach (var item in items.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object || String(item, "uri") is not { Length: > 0 } uri)
            {
                continue;
            }

            var artists = item.TryGetProperty("artists", out var a) && a.ValueKind == JsonValueKind.Array
                ? string.Join(", ", a.EnumerateArray().Select(x => String(x, "name")).Where(n => n.Length > 0))
                : "";
            var album = item.TryGetProperty("album", out var al) && al.ValueKind == JsonValueKind.Object ? al : default;
            results.Add(new SpotifyTrack(
                uri,
                String(item, "name"),
                artists,
                album.ValueKind == JsonValueKind.Object ? String(album, "name") : "",
                album.ValueKind == JsonValueKind.Object ? SmallestImage(album) : null,
                item.TryGetProperty("duration_ms", out var d) && d.TryGetInt64(out var ms) ? ms : 0));
        }

        return results;
    }

    /// <returns>What's playing (null when nothing is), and whether Spotify could be asked at all.</returns>
    private async Task<(NowPlayingInfo? Info, bool Ok)> FetchNowPlayingAsync()
    {
        var response = await CallAsync(HttpMethod.Get, "me/player?additional_types=track,episode").ConfigureAwait(false);
        if (response is null)
        {
            return (null, false);
        }

        using (response)
        {
            if (response.StatusCode == HttpStatusCode.NoContent)
            {
                LastError = null;
                return (null, true); // no active device
            }

            if (!response.IsSuccessStatusCode)
            {
                LastError = await DescribeFailureAsync(response).ConfigureAwait(false);
                return (null, false);
            }

            try
            {
                var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                LastError = null;
                return (ParsePlayback(body), true);
            }
            catch (Exception ex) when (ex is JsonException or InvalidOperationException or KeyNotFoundException)
            {
                LastError = "Spotify sent something unexpected.";
                return (null, false);
            }
        }
    }

    /// <summary>Reads GET /me/player. Null when no track or episode is loaded (nothing at all, or an ad).</summary>
    public static NowPlayingInfo? ParsePlayback(string json)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            return null;
        }

        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        if (!root.TryGetProperty("item", out var item) || item.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        var isPlaying = root.TryGetProperty("is_playing", out var playing) && playing.ValueKind == JsonValueKind.True;
        var progress = root.TryGetProperty("progress_ms", out var p) && p.TryGetInt64(out var ms) ? ms : 0;
        var device = root.TryGetProperty("device", out var d) && d.ValueKind == JsonValueKind.Object && d.TryGetProperty("name", out var dn) ? dn.GetString() : null;

        var title = String(item, "name");
        var duration = item.TryGetProperty("duration_ms", out var dur) && dur.TryGetInt64(out var dms) ? dms : 0;
        var uri = String(item, "uri");

        string artist, album;
        string? artwork;
        if (String(item, "type") == "episode")
        {
            var show = item.TryGetProperty("show", out var s) && s.ValueKind == JsonValueKind.Object ? s : default;
            artist = show.ValueKind == JsonValueKind.Object ? String(show, "name") : "";
            album = artist;
            artwork = LargestImage(item) ?? (show.ValueKind == JsonValueKind.Object ? LargestImage(show) : null);
        }
        else
        {
            artist = item.TryGetProperty("artists", out var artists) && artists.ValueKind == JsonValueKind.Array
                ? string.Join(", ", artists.EnumerateArray().Select(a => String(a, "name")).Where(n => n.Length > 0))
                : "";
            var albumElement = item.TryGetProperty("album", out var a) && a.ValueKind == JsonValueKind.Object ? a : default;
            album = albumElement.ValueKind == JsonValueKind.Object ? String(albumElement, "name") : "";
            artwork = albumElement.ValueKind == JsonValueKind.Object ? LargestImage(albumElement) : null;
        }

        return title.Length == 0
            ? null
            : new NowPlayingInfo(title, artist, ServiceName)
            {
                Album = album,
                ArtworkUrl = artwork,
                IsPlaying = isPlaying,
                Device = device,
                ProgressMs = progress,
                DurationMs = duration,
                Uri = uri,
            };
    }

    /// <summary>
    /// Like <see cref="SendAsync"/>, but when Spotify says there's no active device (nothing
    /// playing anywhere), picks one of the user's open Spotify apps and sends it there instead.
    /// </summary>
    private async Task<bool> SendToDeviceAsync(HttpMethod method, string path, string? json)
    {
        using var response = await CallAsync(method, path, json).ConfigureAwait(false);
        if (response is null)
        {
            return false;
        }

        _nowPlayingCache = null;
        NoDeviceOpen = false;
        if (response.IsSuccessStatusCode)
        {
            LastError = null;
            return true;
        }

        if (response.StatusCode == HttpStatusCode.NotFound && await FindDeviceAsync().ConfigureAwait(false) is { } deviceId)
        {
            var separator = path.Contains('?') ? '&' : '?';
            using var retry = await CallAsync(method, $"{path}{separator}device_id={Uri.EscapeDataString(deviceId)}", json).ConfigureAwait(false);
            if (retry is null)
            {
                return false;
            }

            if (retry.IsSuccessStatusCode)
            {
                LastError = null;
                return true;
            }

            LastError = await DescribeFailureAsync(retry).ConfigureAwait(false);
            return false;
        }

        NoDeviceOpen = response.StatusCode == HttpStatusCode.NotFound;
        LastError = NoDeviceOpen
            ? "No Spotify app is open to play on — open Spotify on this computer or your phone first."
            : await DescribeFailureAsync(response).ConfigureAwait(false);
        return false;
    }

    /// <summary>One of the user's Spotify devices that can be told to play: the active one, else a computer, else any.</summary>
    private async Task<string?> FindDeviceAsync()
    {
        using var response = await CallAsync(HttpMethod.Get, "me/player/devices").ConfigureAwait(false);
        if (response is not { IsSuccessStatusCode: true })
        {
            return null;
        }

        try
        {
            using var document = JsonDocument.Parse(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
            if (!document.RootElement.TryGetProperty("devices", out var devices) || devices.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            return devices.EnumerateArray()
                .Where(d => String(d, "id").Length > 0 && !(d.TryGetProperty("is_restricted", out var r) && r.ValueKind == JsonValueKind.True))
                .OrderByDescending(d => d.TryGetProperty("is_active", out var active) && active.ValueKind == JsonValueKind.True)
                .ThenByDescending(d => String(d, "type") == "Computer")
                .Select(d => String(d, "id"))
                .FirstOrDefault();
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private async Task<bool> SendAsync(HttpMethod method, string path)
    {
        using var response = await CallAsync(method, path).ConfigureAwait(false);
        if (response is null)
        {
            return false;
        }

        _nowPlayingCache = null;
        if (response.IsSuccessStatusCode)
        {
            LastError = null;
            return true;
        }

        LastError = await DescribeFailureAsync(response).ConfigureAwait(false);
        return false;
    }

    /// <summary>Makes one API call with a current access token, renewing it and trying once more if Spotify says it's no longer valid. Null (with LastError set) when Spotify couldn't be asked.</summary>
    private async Task<HttpResponseMessage?> CallAsync(HttpMethod method, string path, string? json = null)
    {
        for (var attempt = 0; attempt < 2; attempt++)
        {
            string accessToken;
            try
            {
                accessToken = await GetAccessTokenAsync(forceRefresh: attempt > 0).ConfigureAwait(false);
            }
            catch (SpotifyAuthException ex)
            {
                SignInExpired |= ex.SignInExpired;
                LastError = ex.SignInExpired ? "Spotify sign-in has expired — reconnect Spotify in the Music tab." : ex.Message;
                return null;
            }

            using var request = new HttpRequestMessage(method, ApiBase + path);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            if (json is not null)
            {
                request.Content = new StringContent(json, Encoding.UTF8, "application/json");
            }
            else if (method != HttpMethod.Get)
            {
                request.Content = new ByteArrayContent([]);
            }

            HttpResponseMessage response;
            try
            {
                response = await _http.SendAsync(request).ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
            {
                LastError = "Couldn't reach Spotify.";
                return null;
            }

            if (response.StatusCode == HttpStatusCode.Unauthorized && attempt == 0)
            {
                response.Dispose();
                continue;
            }

            return response;
        }

        return null;
    }

    private async Task<string> GetAccessTokenAsync(bool forceRefresh)
    {
        await _tokenLock.WaitAsync().ConfigureAwait(false);
        try
        {
            if (!forceRefresh && _tokens is { } tokens && tokens.ExpiresAt - DateTimeOffset.UtcNow > TimeSpan.FromMinutes(1))
            {
                return tokens.AccessToken;
            }

            var fresh = await SpotifyAuthorization.RefreshAsync(_http, _clientId, _refreshToken).ConfigureAwait(false);
            _tokens = fresh;
            if (fresh.RefreshToken != _refreshToken)
            {
                _refreshToken = fresh.RefreshToken;
                _saveRefreshToken(fresh.RefreshToken);
            }

            return fresh.AccessToken;
        }
        finally
        {
            _tokenLock.Release();
        }
    }

    /// <summary>A failed call in words a chat reply can use.</summary>
    private static async Task<string> DescribeFailureAsync(HttpResponseMessage response)
    {
        string? message = null;
        try
        {
            using var document = JsonDocument.Parse(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
            if (document.RootElement.TryGetProperty("error", out var error) && error.ValueKind == JsonValueKind.Object &&
                error.TryGetProperty("message", out var m))
            {
                message = m.GetString();
            }
        }
        catch (JsonException)
        {
        }

        return response.StatusCode switch
        {
            HttpStatusCode.NotFound => "Spotify has no active device — start playing in the Spotify app first.",
            HttpStatusCode.Forbidden when message?.Contains("premium", StringComparison.OrdinalIgnoreCase) == true =>
                "Controlling playback needs Spotify Premium.",
            HttpStatusCode.Forbidden => $"Spotify refused that{(message is null ? "" : $": {message}")}.",
            HttpStatusCode.TooManyRequests => "Spotify is rate-limiting requests — try again in a moment.",
            HttpStatusCode.Unauthorized => "Spotify sign-in isn't working — reconnect Spotify in the Music tab.",
            _ => $"Spotify returned an error ({(int)response.StatusCode}{(message is null ? "" : $": {message}")}).",
        };
    }

    private static string String(JsonElement element, string property) =>
        element.TryGetProperty(property, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() ?? "" : "";

    /// <summary>The smallest image that's still at least thumbnail size (64px), for search results.</summary>
    private static string? SmallestImage(JsonElement element) =>
        element.TryGetProperty("images", out var images) && images.ValueKind == JsonValueKind.Array
            ? images.EnumerateArray()
                .Select(i => (Url: String(i, "url"), Width: i.TryGetProperty("width", out var w) && w.TryGetInt32(out var width) ? width : 0))
                .Where(i => i.Url.Length > 0)
                .OrderBy(i => i.Width >= 64 ? i.Width : int.MaxValue)
                .Select(i => i.Url)
                .FirstOrDefault()
            : null;

    /// <summary>Spotify lists images largest first; take the widest regardless.</summary>
    private static string? LargestImage(JsonElement element) =>
        element.TryGetProperty("images", out var images) && images.ValueKind == JsonValueKind.Array
            ? images.EnumerateArray()
                .OrderByDescending(i => i.TryGetProperty("width", out var w) && w.TryGetInt32(out var width) ? width : 0)
                .Select(i => String(i, "url"))
                .FirstOrDefault(u => u.Length > 0)
            : null;
}
