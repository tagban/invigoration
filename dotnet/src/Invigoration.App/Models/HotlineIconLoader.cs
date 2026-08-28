using Avalonia.Media.Imaging;
using Invigoration.Core.Config;

namespace Invigoration.App.Models;

/// <summary>
/// Hotline user icons aren't a fixed catalog like the Battle.net product icons (GameIconLoader) —
/// a user's icon number is whatever they set on their own real Hotline client, arriving live in
/// chat/user-list packets, so there's no bundling them ahead of time. Fetched on demand from the
/// user's own hlwiki.com icon archive (https://hlwiki.com/ik0ns/{iconId}.png, confirmed live) and
/// cached to disk so the same icon number isn't re-downloaded every session.
/// </summary>
public static class HotlineIconLoader
{
    private static readonly HttpClient Http = new();
    private static readonly Dictionary<ushort, Bitmap?> MemoryCache = [];
    private static readonly Lock SyncRoot = new();

    private static string CacheDirectory => Path.Combine(ConfigStore.DefaultConfigDirectory(), "HotlineIconCache");

    public static async Task<Bitmap?> GetAsync(ushort iconId, CancellationToken ct = default)
    {
        lock (SyncRoot)
        {
            if (MemoryCache.TryGetValue(iconId, out var cached))
            {
                return cached;
            }
        }

        var bitmap = TryLoadFromDisk(iconId) ?? await TryFetchAsync(iconId, ct).ConfigureAwait(false);
        lock (SyncRoot)
        {
            MemoryCache[iconId] = bitmap;
        }

        return bitmap;
    }

    private static Bitmap? TryLoadFromDisk(ushort iconId)
    {
        var path = CachePath(iconId);
        if (!File.Exists(path))
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
            return null;
        }
    }

    private static async Task<Bitmap?> TryFetchAsync(ushort iconId, CancellationToken ct)
    {
        try
        {
            var bytes = await Http.GetByteArrayAsync($"https://hlwiki.com/ik0ns/{iconId}.png", ct).ConfigureAwait(false);
            try
            {
                Directory.CreateDirectory(CacheDirectory);
                await File.WriteAllBytesAsync(CachePath(iconId), bytes, ct).ConfigureAwait(false);
            }
            catch (IOException)
            {
                // Best-effort disk cache — a failed write just means this icon gets re-fetched
                // next time, not a real problem.
            }

            using var stream = new MemoryStream(bytes);
            return new Bitmap(stream);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or NotSupportedException)
        {
            // No icon for this number, or the site's unreachable — a missing icon is a normal,
            // expected case (not every icon number a server assigns has art on hlwiki.com), so
            // this silently returns null rather than surfacing an error for it.
            return null;
        }
    }

    private static string CachePath(ushort iconId) => Path.Combine(CacheDirectory, $"{iconId}.png");
}
