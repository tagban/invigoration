using System.Text.Json;
using System.Text.Json.Serialization;
using Invigoration.Core.Config;

namespace Invigoration.Core.Music;

/// <summary>
/// The music settings — global, not per-bot (one shared Spotify connection, see
/// MusicPlayerRegistry): whether the Music tab and player bar show, and the Spotify connection
/// (the user's own app Client ID, and the saved sign-in). The refresh token is obfuscated at rest
/// the same way bot passwords are; it's a sign-in, so it never goes anywhere but Spotify.
/// </summary>
public static class MusicSettingsStore
{
    private static readonly Lock SyncRoot = new();
    private static StoredSettings? _cached;

    /// <summary>Test hook: points the store at a scratch folder.</summary>
    public static string? DirectoryOverride { get; set; }

    private static string FilePath => Path.Combine(DirectoryOverride ?? ConfigStore.DefaultConfigDirectory(), "music-settings.json");

    /// <summary>Whether the Music tab shows at all — on by default (discoverability), toggled via the Customize menu.</summary>
    public static bool IsEnabled
    {
        get => Current.IsEnabled;
        set => Update(s => s with { IsEnabled = value });
    }

    /// <summary>A thin playback bar docked at the bottom of the whole window, visible whichever tab is showing — opt-in.</summary>
    public static bool ShowBottomBar
    {
        get => Current.ShowBottomBar;
        set => Update(s => s with { ShowBottomBar = value });
    }

    /// <summary>The Client ID of the user's own Spotify developer app.</summary>
    public static string SpotifyClientId
    {
        get => Current.SpotifyClientId;
        set => Update(s => s with { SpotifyClientId = value.Trim() });
    }

    /// <summary>The saved Spotify sign-in, or "" when not connected.</summary>
    public static string SpotifyRefreshToken
    {
        get => Current.SpotifyRefreshToken;
        set => Update(s => s with { SpotifyRefreshToken = value });
    }

    /// <summary>Whether the user has answered the first-run "use Spotify?" question (MainWindow) — a Yes or a No is remembered; closing the question without answering isn't.</summary>
    public static bool SpotifyPromptAnswered
    {
        get => Current.SpotifyPromptAnswered;
        set => Update(s => s with { SpotifyPromptAnswered = value });
    }

    /// <summary>Test hook: forget what's loaded.</summary>
    public static void ResetCacheForTests()
    {
        lock (SyncRoot)
        {
            _cached = null;
        }
    }

    private static StoredSettings Current
    {
        get
        {
            lock (SyncRoot)
            {
                return _cached ??= Load();
            }
        }
    }

    private static void Update(Func<StoredSettings, StoredSettings> change)
    {
        lock (SyncRoot)
        {
            var settings = change(_cached ??= Load());
            _cached = settings;
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
                File.WriteAllText(FilePath, JsonSerializer.Serialize(settings));
            }
            catch (IOException)
            {
                // Best-effort — the in-memory copy still has the change for the rest of this run.
            }
        }
    }

    private static StoredSettings Load()
    {
        try
        {
            // Older files also carry a SelectedService from the retired embedded web player; it's ignored.
            return File.Exists(FilePath)
                ? JsonSerializer.Deserialize<StoredSettings>(File.ReadAllText(FilePath)) ?? new StoredSettings()
                : new StoredSettings();
        }
        catch (Exception ex) when (ex is IOException or JsonException)
        {
            return new StoredSettings();
        }
    }

    private sealed record StoredSettings(
        bool IsEnabled = true,
        bool ShowBottomBar = false,
        string SpotifyClientId = "",
        [property: JsonConverter(typeof(ObfuscatedPasswordJsonConverter))] string SpotifyRefreshToken = "",
        bool SpotifyPromptAnswered = false);
}
