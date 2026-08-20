using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Invigoration.Core.Config;

namespace Invigoration.App.Models;

/// <summary>
/// Loads and caches the chat icons: a user override from
/// <see cref="IconOverrideStore"/> if one exists for the key, otherwise the
/// bundled default under Assets/GameIcons (mostly the original
/// classic.battle.net 28x14 set, with the moderator/channel-op badge swapped
/// for a custom transparent-background green gavel that reads cleanly on the
/// app's dark theme).
/// </summary>
public static class GameIconLoader
{
    private static readonly Dictionary<string, Bitmap?> Cache = [];

    static GameIconLoader() => IconOverrideStore.OverridesChanged += key => Cache.Remove(key);

    public static Bitmap? Get(string key)
    {
        if (string.IsNullOrEmpty(key))
        {
            return null;
        }

        if (Cache.TryGetValue(key, out var cached))
        {
            return cached;
        }

        var bitmap = TryLoadOverride(key) ?? TryLoad($"{key}.png") ?? TryLoad($"{key}.gif") ?? TryLoad($"{key}.jpg");
        Cache[key] = bitmap;
        return bitmap;
    }

    private static Bitmap? TryLoadOverride(string key)
    {
        var path = IconOverrideStore.GetOverridePath(key);
        if (path is null)
        {
            return null;
        }

        try
        {
            using var stream = File.OpenRead(path);
            return new Bitmap(stream);
        }
        catch (Exception ex) when (ex is IOException or NotSupportedException)
        {
            // Corrupt/unreadable override file — fall back to the bundled default below.
            return null;
        }
    }

    private static Bitmap? TryLoad(string fileName)
    {
        try
        {
            var uri = new Uri($"avares://Invigoration.App/Assets/GameIcons/{fileName}");
            using var stream = AssetLoader.Open(uri);
            return new Bitmap(stream);
        }
        catch (FileNotFoundException)
        {
            return null;
        }
    }
}
