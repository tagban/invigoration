using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Invigoration.Core.Music.Pandora;

public sealed record PandoraStation(string StationToken, string StationName);

public sealed record PandoraAudioStream(string AudioUrl, string Bitrate, string Encoding);

/// <summary>An actual song. Ad items are filtered out by <see cref="PandoraApiClient.GetPlaylistAsync"/> before they ever become one of these — see its remarks. SongRating is 1 for thumbs-up, otherwise unrated (Pandora sends this as a JSON number, confirmed live — a naive string cast throws).</summary>
public sealed record PandoraTrack(string TrackToken, string SongName, string ArtistName, string AlbumName, int? SongRating, PandoraAudioStream? AudioStream);

/// <summary>A Pandora API call failed at the protocol level (bad credentials, station gone, etc.) — <see cref="Code"/> is the numeric code straight from the JSON envelope (1011/1012 = bad username/password; see PandoraApiClient's remarks for the others this client handles specially).</summary>
public sealed class PandoraApiException(string message, int code) : Exception(message)
{
    public int Code { get; } = code;
}

/// <summary>
/// A thin C# port of pydora's (github.com/mcrute/pydora) real client — the same unofficial,
/// long-documented JSON-RPC API Pandora's own official apps use (see 6xq.net/pandora-apidoc/),
/// confirmed against pydora's actual source rather than the docs page alone (which had gaps: it
/// doesn't mention that the decrypted syncTime is an ASCII decimal string, not a raw binary
/// integer, for instance). Talking to this directly — rather than automating a web page like the
/// other services (see WebViewMusicController) — sidesteps Pandora's web player having no stable,
/// documented DOM to script against and needing a real audio-playback pipeline in the app instead
/// (see PandoraPlayerController in Invigoration.App).
/// </summary>
public sealed class PandoraApiClient(HttpClient? httpClient = null) : IDisposable
{
    private const string ApiHost = "tuner.pandora.com";
    private const string ApiVersion = "5";

    private static readonly string[] TlsRequiredMethods =
    [
        "auth.partnerLogin", "auth.userLogin", "station.getPlaylist", "user.createUser",
    ];

    private readonly HttpClient _http = httpClient ?? new HttpClient();
    private readonly bool _ownsHttpClient = httpClient is null;

    private string? _partnerAuthToken;
    private string? _partnerId;
    private string? _userAuthToken;
    private string? _userId;
    private long _serverSyncTimeAtLogin;
    private long _loginTimestamp;
    private string? _username;
    private string? _password;

    public bool IsLoggedIn => _userAuthToken is not null;

    /// <summary>Full login: partner login (identifies us as the Android app) then user login (the real account) — pydora's <c>_authenticate</c>. Throws <see cref="PandoraApiException"/> with code 1011/1012 for a bad username/password so callers can show a real error instead of a generic one.</summary>
    public async Task LoginAsync(string username, string password, CancellationToken ct = default)
    {
        _username = username;
        _password = password;
        await PartnerLoginAsync(ct).ConfigureAwait(false);
        await UserLoginAsync(ct).ConfigureAwait(false);
    }

    private async Task PartnerLoginAsync(CancellationToken ct)
    {
        var body = new JsonObject
        {
            ["username"] = PandoraPartnerCredentials.Username,
            ["password"] = PandoraPartnerCredentials.Password,
            ["deviceModel"] = PandoraPartnerCredentials.DeviceModel,
            ["version"] = ApiVersion,
        };

        var result = await CallAsync("auth.partnerLogin", body, encrypt: false, ct).ConfigureAwait(false);
        _partnerAuthToken = AsString(result["partnerAuthToken"]);
        _partnerId = AsString(result["partnerId"]);

        var syncTimeHex = AsString(result["syncTime"]) ?? throw new PandoraApiException("Partner login response had no syncTime", 0);
        var decrypted = PandoraCrypto.DecryptRaw(PandoraPartnerCredentials.DecryptKey, syncTimeHex);
        _serverSyncTimeAtLogin = PandoraCrypto.ParseSyncTimeDigits(decrypted);
        _loginTimestamp = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
    }

    private async Task UserLoginAsync(CancellationToken ct)
    {
        var body = new JsonObject
        {
            ["loginType"] = "user",
            ["username"] = _username,
            ["password"] = _password,
            ["partnerAuthToken"] = _partnerAuthToken,
            ["includePandoraOneInfo"] = true,
        };

        var result = await CallAsync("auth.userLogin", body, encrypt: true, ct).ConfigureAwait(false);
        _userAuthToken = AsString(result["userAuthToken"]);
        _userId = AsString(result["userId"]);
    }

    public async Task<IReadOnlyList<PandoraStation>> GetStationListAsync(CancellationToken ct = default)
    {
        var body = new JsonObject { ["includeStationArtUrl"] = true };
        var result = await AuthenticatedCallAsync("user.getStationList", body, ct).ConfigureAwait(false);
        var stations = result["stations"]?.AsArray() ?? [];
        return stations
            .Select(s => new PandoraStation(AsString(s!["stationToken"]) ?? "", AsString(s["stationName"]) ?? ""))
            .ToList();
    }

    /// <summary>
    /// Ad items (an <c>items[]</c> entry with a populated <c>adToken</c> instead of a
    /// <c>trackToken</c> — pydora's <c>Track.is_ad</c>) are silently dropped rather than surfaced:
    /// playing real ad audio isn't the point of this integration and pydora's own ad-fetch path
    /// adds real complexity (a separate <c>ad.getAdMetadata</c> call) for no benefit here — a
    /// known, accepted gap, not an oversight.
    /// </summary>
    public async Task<IReadOnlyList<PandoraTrack>> GetPlaylistAsync(string stationToken, CancellationToken ct = default)
    {
        var body = new JsonObject
        {
            ["stationToken"] = stationToken,
            ["includeTrackLength"] = true,
            ["xplatformAdCapable"] = false,
        };

        var result = await AuthenticatedCallAsync("station.getPlaylist", body, ct).ConfigureAwait(false);
        var items = result["items"]?.AsArray() ?? [];
        var tracks = new List<PandoraTrack>();
        foreach (var item in items)
        {
            if (item is null || item["adToken"] is not null || item["trackToken"] is null)
            {
                continue;
            }

            var audioMap = item["audioUrlMap"];
            var high = audioMap?["highQuality"];
            var audio = high is null
                ? null
                : new PandoraAudioStream(AsString(high["audioUrl"]) ?? "", AsString(high["bitrate"]) ?? "", AsString(high["encoding"]) ?? "");

            tracks.Add(new PandoraTrack(
                AsString(item["trackToken"]) ?? "",
                AsString(item["songName"]) ?? "",
                AsString(item["artistName"]) ?? "",
                AsString(item["albumName"]) ?? "",
                AsInt(item["songRating"]),
                audio));
        }

        return tracks;
    }

    public async Task<bool> AddFeedbackAsync(string trackToken, bool isPositive, CancellationToken ct = default)
    {
        var body = new JsonObject { ["trackToken"] = trackToken, ["isPositive"] = isPositive };
        try
        {
            await AuthenticatedCallAsync("station.addFeedback", body, ct).ConfigureAwait(false);
            return true;
        }
        catch (PandoraApiException)
        {
            return false;
        }
    }

    /// <summary>Adds auth to the body/query and retries once on code 1001 (InvalidAuthToken) by re-running the full login — pydora's <c>BaseAPIClient.__call__</c>: <c>except InvalidAuthToken: self._authenticate(); return self.transport(...)</c>.</summary>
    private async Task<JsonObject> AuthenticatedCallAsync(string method, JsonObject body, CancellationToken ct)
    {
        try
        {
            return await CallAsync(method, body, encrypt: true, ct, authenticated: true).ConfigureAwait(false);
        }
        catch (PandoraApiException ex) when (ex.Code == 1001 && _username is not null && _password is not null)
        {
            await LoginAsync(_username, _password, ct).ConfigureAwait(false);
            return await CallAsync(method, body, encrypt: true, ct, authenticated: true).ConfigureAwait(false);
        }
    }

    private async Task<JsonObject> CallAsync(string method, JsonObject body, bool encrypt, CancellationToken ct, bool authenticated = false)
    {
        if (authenticated)
        {
            body["userAuthToken"] = _userAuthToken;
            body["syncTime"] = CurrentSyncTime;
        }
        else if (method == "auth.userLogin")
        {
            body["syncTime"] = CurrentSyncTime;
        }

        var scheme = TlsRequiredMethods.Contains(method) ? "https" : "http";
        var query = BuildQuery(method, authenticated);
        var uri = $"{scheme}://{ApiHost}/services/json/?{query}";

        var payload = encrypt
            ? PandoraCrypto.Encrypt(PandoraPartnerCredentials.EncryptKey, body.ToJsonString())
            : body.ToJsonString();

        using var response = await _http.PostAsync(uri, new StringContent(payload, Encoding.UTF8, "text/plain"), ct).ConfigureAwait(false);
        var responseText = await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
        var envelope = JsonNode.Parse(responseText)?.AsObject() ?? throw new PandoraApiException("Empty/invalid response from Pandora", 0);

        var stat = AsString(envelope["stat"]);
        if (stat != "ok")
        {
            var code = AsInt(envelope["code"]) ?? 0;
            var message = AsString(envelope["message"]) ?? "Pandora API call failed";
            throw new PandoraApiException(message, code);
        }

        return envelope["result"]?.AsObject() ?? [];
    }

    private string BuildQuery(string method, bool authenticated)
    {
        var pairs = new List<string> { $"method={method}" };
        if (authenticated)
        {
            pairs.Add($"auth_token={Uri.EscapeDataString(_userAuthToken ?? "")}");
            pairs.Add($"partner_id={Uri.EscapeDataString(_partnerId ?? "")}");
            pairs.Add($"user_id={Uri.EscapeDataString(_userId ?? "")}");
        }
        else if (method == "auth.userLogin")
        {
            pairs.Add($"auth_token={Uri.EscapeDataString(_partnerAuthToken ?? "")}");
            pairs.Add($"partner_id={Uri.EscapeDataString(_partnerId ?? "")}");
        }

        return string.Join('&', pairs);
    }

    private long CurrentSyncTime => _serverSyncTimeAtLogin + (DateTimeOffset.UtcNow.ToUnixTimeSeconds() - _loginTimestamp);

    /// <summary>
    /// Reads a JsonNode as a string regardless of its actual JSON kind — confirmed live that
    /// Pandora sends some fields (songRating, at least) as a JSON number rather than the string
    /// pydora's docs implied, and a direct <c>(string?)node</c> cast throws
    /// InvalidOperationException on those instead of just... being the value. TryGetValue never
    /// throws, so this degrades gracefully instead of taking down a whole playlist fetch over one
    /// unexpectedly-typed field.
    /// </summary>
    private static string? AsString(JsonNode? node)
    {
        if (node is not JsonValue value)
        {
            return null;
        }

        if (value.TryGetValue<string>(out var s))
        {
            return s;
        }

        return value.ToJsonString();
    }

    private static int? AsInt(JsonNode? node)
    {
        if (node is not JsonValue value)
        {
            return null;
        }

        if (value.TryGetValue<int>(out var i))
        {
            return i;
        }

        return int.TryParse(AsString(node), out var parsed) ? parsed : null;
    }

    public void Dispose()
    {
        if (_ownsHttpClient)
        {
            _http.Dispose();
        }
    }
}
