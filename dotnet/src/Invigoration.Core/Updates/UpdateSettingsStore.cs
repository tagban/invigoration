using System.Text.Json;
using Invigoration.Core.Config;

namespace Invigoration.Core.Updates;

/// <summary>
/// Whether to look for new releases, and which version the user has already waved away — global,
/// not per-bot. Deliberately tiny: this feature never downloads or installs anything, it only
/// says "there's a newer one" once and links to the releases page.
/// </summary>
public static class UpdateSettingsStore
{
    private static readonly Lock SyncRoot = new();
    private static StoredSettings? _cached;

    /// <summary>Test hook: points the store at a scratch folder.</summary>
    public static string? DirectoryOverride { get; set; }

    private static string FilePath => Path.Combine(DirectoryOverride ?? ConfigStore.DefaultConfigDirectory(), "update-settings.json");

    /// <summary>
    /// Whether to check GitHub for a newer release on startup — on by default, since the notice is
    /// a single dismissible line and most people want to know. Turning it off stops the app
    /// contacting GitHub at all.
    /// </summary>
    public static bool CheckForUpdates
    {
        get => Current.CheckForUpdates;
        set => Update(s => s with { CheckForUpdates = value });
    }

    /// <summary>
    /// The newest version the user has dismissed, or "". Dismissing hides that version for good;
    /// a later release is newer than what's stored here, so it shows up on its own.
    /// </summary>
    public static string DismissedVersion
    {
        get => Current.DismissedVersion;
        set => Update(s => s with { DismissedVersion = value });
    }

    /// <summary>Whether <paramref name="version"/> should be shown, given what's already been dismissed.</summary>
    public static bool ShouldAnnounce(ReleaseVersion version) =>
        ReleaseVersion.TryParse(DismissedVersion) is not { } dismissed || version.IsNewerThan(dismissed);

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
            return File.Exists(FilePath)
                ? JsonSerializer.Deserialize<StoredSettings>(File.ReadAllText(FilePath)) ?? new StoredSettings()
                : new StoredSettings();
        }
        catch (Exception ex) when (ex is IOException or JsonException)
        {
            return new StoredSettings();
        }
    }

    private sealed record StoredSettings(bool CheckForUpdates = true, string DismissedVersion = "");
}
