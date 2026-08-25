using System.Text.Json;

namespace Invigoration.Core.Music;

/// <summary>
/// Which music service the embedded player tab remembers across restarts — global, not per-bot
/// (same reasoning as MusicPlayerRegistry: one shared player for the whole app). A single small
/// JSON file rather than IconOverrideStore's folder-of-files approach, since there's only one
/// value to persist.
/// </summary>
public static class MusicSettingsStore
{
    private static string FilePath => Path.Combine(Config.ConfigStore.DefaultConfigDirectory(), "music-settings.json");

    private static StoredSettings? _cached;

    public static MusicService SelectedService
    {
        get => Current.SelectedService;
        set => Save(Current with { SelectedService = value });
    }

    /// <summary>Whether the Music tab shows at all — off by default is wrong here (the whole point is discoverability), on by default, toggled via the Customize menu for anyone who doesn't want it.</summary>
    public static bool IsEnabled
    {
        get => Current.IsEnabled;
        set => Save(Current with { IsEnabled = value });
    }

    /// <summary>A thin persistent playback-control bar docked at the bottom of the whole window, visible no matter which top-level tab is showing — explicitly opt-in (off by default), for controlling playback without needing to switch to the Music tab.</summary>
    public static bool ShowBottomBar
    {
        get => Current.ShowBottomBar;
        set => Save(Current with { ShowBottomBar = value });
    }

    private static StoredSettings Current
    {
        get
        {
            if (_cached is { } cached)
            {
                return cached;
            }

            var loaded = Load();
            _cached = loaded;
            return loaded;
        }
    }

    private static void Save(StoredSettings settings)
    {
        _cached = settings;
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(settings));
        }
        catch (IOException)
        {
            // Best-effort — the in-memory cache still reflects the change for the rest of this
            // run even if the write itself failed.
        }
    }

    private static StoredSettings Load()
    {
        try
        {
            if (!File.Exists(FilePath))
            {
                return new StoredSettings(MusicService.YouTubeMusic, true, false);
            }

            var stored = JsonSerializer.Deserialize<StoredSettings>(File.ReadAllText(FilePath));
            return stored ?? new StoredSettings(MusicService.YouTubeMusic, true, false);
        }
        catch (Exception ex) when (ex is IOException or JsonException)
        {
            return new StoredSettings(MusicService.YouTubeMusic, true, false);
        }
    }

    private sealed record StoredSettings(MusicService SelectedService, bool IsEnabled, bool ShowBottomBar);
}
