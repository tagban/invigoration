using Avalonia;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Invigoration.Core.Config;
using Invigoration.Core.StatString;

namespace Invigoration.App.Models;

/// <summary>
/// Diablo II realm characters drawn as themselves — in their own gear, idling the way the character
/// select screen shows them — from the character pack the user downloaded (D2EquipmentStore's
/// d2-characters.zip, composed by Core's D2CharacterPack). Only for the character dock, and only
/// once a pack is stored; everyone else keeps their chat avatar (ChatAvatarLoader).
/// </summary>
/// <remarks>
/// Composed figures are shared by everyone with the same 33-byte portrait (same class, gear, tints
/// and stance), and the cache is bounded, since a mass join can bring thousands of characters. The
/// pack is read into memory rather than held open, so an update can replace the file on any OS;
/// a newly saved pack is picked up on the next request.
/// </remarks>
public static class D2CharacterLoader
{
    /// <summary>Clear ground under the feet, matching where the Arreat Summit avatars stand in their 88-pixel frames.</summary>
    public const int GroundGap = 7;

    private const int MaxCached = 128;

    private static readonly Dictionary<string, SpriteAnimation?> Cache = new(StringComparer.Ordinal);
    private static D2CharacterPack? _pack;
    private static int _stale = 1;

    static D2CharacterLoader() => D2EquipmentStore.Changed += () => Interlocked.Exchange(ref _stale, 1);

    /// <summary>The dressed, animated figure for a statstring, or null when there's no pack or nothing it can draw (another game, an Open or dead hardcore character). UI thread only.</summary>
    public static SpriteAnimation? Get(string statString)
    {
        if (D2Character.PortraitBytes(statString) is not { } portrait)
        {
            return null;
        }

        if (Interlocked.Exchange(ref _stale, 0) == 1)
        {
            Reload();
        }

        if (_pack is not { } pack)
        {
            return null;
        }

        var key = Convert.ToHexString(portrait);
        if (Cache.TryGetValue(key, out var cached))
        {
            return cached;
        }

        var animation = pack.Compose(portrait) is { } frames ? ToAnimation(frames) : null;
        if (Cache.Count >= MaxCached)
        {
            Cache.Clear();
        }

        Cache[key] = animation;
        return animation;
    }

    private static void Reload()
    {
        Cache.Clear();
        _pack?.Dispose();
        _pack = null;
        if (!D2EquipmentStore.HasCharacterPack)
        {
            return;
        }

        try
        {
            _pack = D2CharacterPack.Open(new MemoryStream(File.ReadAllBytes(D2EquipmentStore.CharacterPackPath)));
        }
        catch (Exception ex) when (ex is FormatException or IOException or UnauthorizedAccessException)
        {
            // An unreadable pack draws nobody; the Unknown avatar stands in until a good one arrives.
            _pack = null;
        }
    }

    private static SpriteAnimation ToAnimation(D2CharacterFrames frames)
    {
        var stripWidth = frames.FrameWidth * frames.FrameDurationsMs.Count;
        var height = frames.FrameHeight + GroundGap;
        var bitmap = new WriteableBitmap(new PixelSize(stripWidth, height), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);
        using (var framebuffer = bitmap.Lock())
        {
            // RGBA → BGRA. Every pixel is fully opaque or fully clear (and clear ones are zero), so
            // straight and premultiplied alpha are the same bytes.
            var row = new byte[stripWidth * 4];
            for (var y = 0; y < height; y++)
            {
                Array.Clear(row);
                if (y < frames.FrameHeight)
                {
                    var offset = y * stripWidth * 4;
                    for (var i = 0; i < row.Length; i += 4)
                    {
                        row[i] = frames.Strip[offset + i + 2];
                        row[i + 1] = frames.Strip[offset + i + 1];
                        row[i + 2] = frames.Strip[offset + i];
                        row[i + 3] = frames.Strip[offset + i + 3];
                    }
                }

                System.Runtime.InteropServices.Marshal.Copy(row, 0, framebuffer.Address + y * framebuffer.RowBytes, row.Length);
            }
        }

        return new SpriteAnimation(bitmap, frames.FrameWidth, height, frames.FrameDurationsMs);
    }
}
