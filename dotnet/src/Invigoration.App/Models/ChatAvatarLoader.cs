using Avalonia.Media.Imaging;
using Avalonia.Platform;

namespace Invigoration.App.Models;

/// <summary>A looping animation stored as one horizontal strip of equal-sized frames.</summary>
public sealed record SpriteAnimation(Bitmap Sheet, int FrameWidth, int FrameHeight, IReadOnlyList<int> FrameDurationsMs)
{
    public int TotalDurationMs { get; } = FrameDurationsMs.Sum();

    /// <summary>Which frame is showing <paramref name="elapsedMs"/> into the loop.</summary>
    public int FrameAt(long elapsedMs)
    {
        var t = (int)(elapsedMs % TotalDurationMs);
        for (var i = 0; i < FrameDurationsMs.Count; i++)
        {
            t -= FrameDurationsMs[i];
            if (t < 0)
            {
                return i;
            }
        }

        return FrameDurationsMs.Count - 1;
    }
}

/// <summary>
/// Diablo II's Battle.net chat avatars (see Core's ChatAvatar for who gets which) — the animations
/// from The Arreat Summit, bundled under Assets/D2Avatars as strips of 78×88 frames with the black
/// backdrop keyed out and a shared crop, so every figure stands on the same baseline. Frame timings
/// are the originals'.
/// </summary>
public static class ChatAvatarLoader
{
    private const int FrameWidth = 78;
    private const int FrameHeight = 88;

    private static readonly Dictionary<string, int[]> FrameTimings = new()
    {
        ["blizzrep"] = [170, 170, 170, 170, 170, 170],
        ["sysop"] = [150, 150, 150, 150, 150, 150],
        ["moderator"] = [180, 180, 180, 180, 180, 180],
        ["speaker"] = [120, 120, 120, 120, 120, 120, 120, 120],
        ["referee"] = [140, 140, 140, 140, 140, 140, 140, 140],
        ["deadhardcore"] = [180, 180, 180, 180, 180, 180, 180, 180],
        ["unknown"] = [150, 150, 150, 150, 150, 150, 150, 150],
        ["chatclient"] = [160, 160, 160, 160, 160, 160],
        ["starcraft"] = [120, 120, 120, 120, 120, 120, 120, 120],
        ["broodwar"] = [150, 130, 130, 140, 150, 160],
        ["war2"] = [150, 150, 150, 150, 150, 150],
        ["diablo"] = [120, 120, 120, 120, 120, 120, 120, 120],
    };

    private static readonly Dictionary<string, SpriteAnimation?> Cache = [];

    public static SpriteAnimation? Get(string? key)
    {
        if (key is null || !FrameTimings.TryGetValue(key, out var timings))
        {
            return null;
        }

        if (Cache.TryGetValue(key, out var cached))
        {
            return cached;
        }

        SpriteAnimation? animation;
        try
        {
            using var stream = AssetLoader.Open(new Uri($"avares://Invigoration.App/Assets/D2Avatars/{key}.png"));
            var sheet = new Bitmap(stream);
            // A strip that doesn't hold exactly the expected frames would animate garbage — treat it as missing.
            animation = sheet.PixelSize.Width == FrameWidth * timings.Length && sheet.PixelSize.Height == FrameHeight
                ? new SpriteAnimation(sheet, FrameWidth, FrameHeight, timings)
                : null;
        }
        catch (FileNotFoundException)
        {
            animation = null;
        }

        Cache[key] = animation;
        return animation;
    }
}
