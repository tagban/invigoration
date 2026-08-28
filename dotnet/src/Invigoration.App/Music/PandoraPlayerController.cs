using Invigoration.Core.Music;
using Invigoration.Core.Music.Pandora;
using ManagedBass;

namespace Invigoration.App.Music;

/// <summary>
/// Pandora's own web player has no stable, documented DOM to script the way YouTube Music's and
/// Spotify's do (see WebViewMusicController) — this instead talks to Pandora's real JSON-RPC API
/// directly (PandoraApiClient) and plays the MP3 stream URLs it returns through ManagedBass, a
/// real audio-decode-and-playback pipeline the other two services don't need since a WebView
/// already has one built in.
/// </summary>
public sealed class PandoraPlayerController : IMusicPlayerController, IDisposable
{
    private readonly PandoraApiClient _client = new();
    private readonly object _sync = new();

    private List<PandoraTrack> _queue = [];
    private int _index = -1;
    private int _streamHandle;
    private string? _currentStationToken;

    private static bool _bassInitialized;

    public bool IsLoggedIn => _client.IsLoggedIn;

    /// <summary>The currently-playing station's token, if any — lets other UI (the bottom playback bar's station picker, not just the Music tab's own) show which station is already selected without duplicating playback state.</summary>
    public string? CurrentStationToken => _currentStationToken;

    /// <summary>Logs in and, on success, persists the credentials for next launch — mirrors how BotConfig already stores Battle.net credentials plaintext-at-rest (see PandoraCredentialsStore's remarks).</summary>
    public async Task<bool> LoginAsync(string username, string password)
    {
        EnsureBassInitialized();
        try
        {
            await _client.LoginAsync(username, password).ConfigureAwait(false);
            PandoraCredentialsStore.Username = username;
            PandoraCredentialsStore.Password = password;
            return true;
        }
        catch (PandoraApiException)
        {
            return false;
        }
    }

    public Task<IReadOnlyList<PandoraStation>> GetStationsAsync() => _client.GetStationListAsync();

    /// <summary>Starts (or restarts) a station — fetches its first playlist fragment and plays the first track. Subsequent tracks are fetched automatically as the queue empties (see AdvanceAsync).</summary>
    public async Task PlayStationAsync(string stationToken)
    {
        _currentStationToken = stationToken;
        _queue = [.. await _client.GetPlaylistAsync(stationToken).ConfigureAwait(false)];
        _index = -1;
        await AdvanceAsync().ConfigureAwait(false);
    }

    public Task<bool> SkipAsync() => AdvanceAsync();

    public Task<bool> PlayPauseAsync()
    {
        lock (_sync)
        {
            if (_streamHandle == 0)
            {
                return Task.FromResult(false);
            }

            var isPlaying = Bass.ChannelIsActive(_streamHandle) == PlaybackState.Playing;
            var ok = isPlaying ? Bass.ChannelPause(_streamHandle) : Bass.ChannelPlay(_streamHandle);
            return Task.FromResult(ok);
        }
    }

    public Task<bool> ThumbsUpAsync() => RateCurrentAsync(isPositive: true);

    public Task<bool> ThumbsDownAsync() => RateCurrentAsync(isPositive: false);

    private async Task<bool> RateCurrentAsync(bool isPositive)
    {
        var current = CurrentTrack;
        if (current is null)
        {
            return false;
        }

        return await _client.AddFeedbackAsync(current.TrackToken, isPositive).ConfigureAwait(false);
    }

    public Task<NowPlayingInfo?> GetNowPlayingAsync()
    {
        var current = CurrentTrack;
        return Task.FromResult(current is null ? null : new NowPlayingInfo(current.SongName, current.ArtistName, "Pandora"));
    }

    private PandoraTrack? CurrentTrack => _index >= 0 && _index < _queue.Count ? _queue[_index] : null;

    /// <summary>
    /// Stops whatever's playing and moves to the next queued track, fetching a fresh playlist
    /// fragment from the current station once the queue runs dry — pydora's own clients do the
    /// same thing (a "playlist" here is really just a batch of a handful of upcoming tracks, not
    /// the whole station). Also wired up as the BASS "channel ended" callback (see
    /// Bass.ChannelSetSync below) so a station keeps playing continuously without needing a manual
    /// !skip after every song.
    /// </summary>
    private async Task<bool> AdvanceAsync()
    {
        StopCurrentStream();
        _index++;

        if (_index >= _queue.Count)
        {
            if (_currentStationToken is null)
            {
                return false;
            }

            _queue = [.. await _client.GetPlaylistAsync(_currentStationToken).ConfigureAwait(false)];
            _index = 0;
        }

        if (_queue.Count == 0 || CurrentTrack is not { AudioStream: not null } track)
        {
            return false;
        }

        var handle = Bass.CreateStream(track.AudioStream!.AudioUrl, 0, BassFlags.Default, null);
        if (handle == 0)
        {
            return false;
        }

        lock (_sync)
        {
            _streamHandle = handle;
        }

        Bass.ChannelPlay(handle);
        Bass.ChannelSetSync(handle, SyncFlags.End, 0, (_, _, _, _) => _ = AdvanceAsync(), IntPtr.Zero);
        return true;
    }

    private void StopCurrentStream()
    {
        lock (_sync)
        {
            if (_streamHandle != 0)
            {
                Bass.StreamFree(_streamHandle);
                _streamHandle = 0;
            }
        }
    }

    private static void EnsureBassInitialized()
    {
        if (_bassInitialized)
        {
            return;
        }

        // Device -1 = default output device; matches ManagedBass's own recommended default init.
        Bass.Init();
        _bassInitialized = true;
    }

    public void Dispose()
    {
        StopCurrentStream();
        _client.Dispose();
    }
}
