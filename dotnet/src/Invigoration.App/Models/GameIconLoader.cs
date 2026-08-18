using Avalonia.Media.Imaging;
using Avalonia.Platform;

namespace Invigoration.App.Models;

/// <summary>
/// Loads and caches the chat icons bundled under Assets/GameIcons: mostly
/// the original classic.battle.net 28x14 set, with the moderator/channel-op
/// badge swapped for the larger WC3-ladder "General Icons" version, which
/// reads more clearly even scaled down small.
/// </summary>
public static class GameIconLoader
{
    private static readonly Dictionary<string, Bitmap?> Cache = [];

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

        Bitmap? bitmap = TryLoad($"{key}.png") ?? TryLoad($"{key}.gif") ?? TryLoad($"{key}.jpg");
        Cache[key] = bitmap;
        return bitmap;
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
