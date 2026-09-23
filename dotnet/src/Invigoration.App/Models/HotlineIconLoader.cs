using System.IO.Compression;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Invigoration.Core.Config;
using SkiaSharp;

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
    /// <summary>hlwiki.com's own gallery of every icon it has — what the "ik0ns page" links open.</summary>
    public const string GalleryUrl = "https://hlwiki.com/ik0ns/";

    /// <summary>Every icon in one archive, so "Download All" is one request instead of ~6,500. Optional — when it isn't there, the picker falls back to fetching icons one at a time.</summary>
    public const string ZipUrl = "https://hlwiki.com/ik0ns/ik0ns.zip";

    private static readonly HttpClient Http = new();
    private static readonly Dictionary<ushort, Bitmap?> MemoryCache = [];
    private static readonly Lock SyncRoot = new();

    /// <summary>
    /// Where a name starts over its icon, in pixels. Measured across the ~4,200 standard 232x18
    /// banners on hlwiki.com: most carry a small emblem at the left that ends between x=26 and 30,
    /// with x=29 by far the most common hard edge, so 30 clears it. (20 cut through most emblems.)
    /// Then 3 more, a bold lowercase "l" further, to line up with where the banner's own lettering
    /// area starts — checked by eye against a live user list.
    /// </summary>
    public const double NameOffset = 33;

    /// <summary>The Users panel's own background — what a transparent part of an icon shows through to.</summary>
    public static readonly Color PanelBackground = Color.FromRgb(0xD8, 0xD8, 0xD8);

    /// <summary>How far past NameOffset to look when judging the icon's brightness — roughly a typical name's width.</summary>
    private const int NameSampleWidth = 90;

    private static readonly Dictionary<ushort, bool> DarkBehindName = [];

    internal static string CacheDirectory => Path.Combine(ConfigStore.DefaultConfigDirectory(), "HotlineIconCache");

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
            return Decode(iconId, File.ReadAllBytes(path));
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

            return Decode(iconId, bytes);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or NotSupportedException)
        {
            // No icon for this number, or the site's unreachable — a missing icon is a normal,
            // expected case (not every icon number a server assigns has art on hlwiki.com), so
            // this silently returns null rather than surfacing an error for it.
            return null;
        }
    }

    /// <summary>
    /// Whether the part of this icon a name is drawn over is dark enough that the name should be
    /// white rather than black. Known once GetAsync has loaded the icon; false (black) before that
    /// or when there's no icon, which matches the light panel a missing icon leaves showing.
    /// </summary>
    public static bool IsDarkBehindName(ushort iconId)
    {
        lock (SyncRoot)
        {
            return DarkBehindName.TryGetValue(iconId, out var dark) && dark;
        }
    }

    /// <summary>Black or white — whichever reads better over this icon. Call after GetAsync.</summary>
    public static IBrush NameBrushFor(ushort iconId) => IsDarkBehindName(iconId) ? Brushes.White : Brushes.Black;

    private static Bitmap Decode(ushort iconId, byte[] bytes)
    {
        var dark = MeasureDarkBehindName(bytes);
        lock (SyncRoot)
        {
            DarkBehindName[iconId] = dark;
        }

        using var stream = new MemoryStream(bytes);
        return new Bitmap(stream);
    }

    /// <summary>
    /// Averages the brightness of the strip a name covers (from NameOffset, about a name's width,
    /// full height), with transparent pixels showing the panel behind them — a small 16x16 icon
    /// leaves the name entirely over the light panel, so it stays black. Below half brightness the
    /// name goes white.
    ///
    /// Brightness is the perceived (Rec. 601) weighting on plain RGB, not the WCAG relative
    /// luminance this first used: WCAG called ~870 of the banners "light" — saturated mid-blues,
    /// reds and busy photos — and gave them black names that were plainly harder to read than
    /// white ones (a user's dark blue icon with black text is what showed it). Compared side by
    /// side across the whole hlwiki.com set, this rule picked the more readable color on those.
    /// </summary>
    internal static bool MeasureDarkBehindName(byte[] bytes)
    {
        using var bitmap = SKBitmap.Decode(bytes);
        if (bitmap is null)
        {
            return false;
        }

        var start = (int)NameOffset;
        var end = Math.Min(bitmap.Width, start + NameSampleWidth);
        if (end <= start)
        {
            return false;
        }

        var panel = Brightness(PanelBackground.R, PanelBackground.G, PanelBackground.B);
        double total = 0;
        var count = 0;
        for (var y = 0; y < bitmap.Height; y++)
        {
            for (var x = start; x < end; x++)
            {
                var c = bitmap.GetPixel(x, y);
                var a = c.Alpha / 255.0;
                total += Brightness(c.Red, c.Green, c.Blue) * a + panel * (1 - a);
                count++;
            }
        }

        // Pixels in the name's strip past a narrow icon's right edge are panel, too.
        var pastIcon = (start + NameSampleWidth - end) * bitmap.Height;
        total += panel * pastIcon;
        count += pastIcon;

        return total / count < 0.5;

        static double Brightness(byte r, byte g, byte b) => (0.299 * r + 0.587 * g + 0.114 * b) / 255.0;
    }

    private static string CachePath(ushort iconId) => Path.Combine(CacheDirectory, $"{iconId}.png");

    /// <summary>
    /// Fetches ZipUrl and unpacks every <c>NNN.png</c> in it into the icon cache — folders inside
    /// the zip don't matter, only the file name, which also means an entry can't be written
    /// anywhere but the cache. Returns how many icons it saved, or null when there's no zip on
    /// the site (so the caller can fall back to one-at-a-time). Reports progress as text.
    /// </summary>
    public static async Task<int?> DownloadZipAsync(IProgress<string> progress, CancellationToken ct = default)
    {
        using var response = await Http.GetAsync(ZipUrl, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            return null;
        }

        var total = response.Content.Headers.ContentLength;
        var tempPath = Path.Combine(Path.GetTempPath(), $"invigoration-ik0ns-{Guid.NewGuid():N}.zip");
        try
        {
            await using (var body = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false))
            await using (var file = File.Create(tempPath))
            {
                var buffer = new byte[81920];
                long read = 0;
                int n;
                while ((n = await body.ReadAsync(buffer, ct).ConfigureAwait(false)) > 0)
                {
                    await file.WriteAsync(buffer.AsMemory(0, n), ct).ConfigureAwait(false);
                    read += n;
                    progress.Report(total is > 0
                        ? $"Downloading icon pack... {read * 100 / total.Value}% of {total.Value / 1048576.0:0.#} MB"
                        : $"Downloading icon pack... {read / 1048576.0:0.#} MB");
                }
            }

            progress.Report("Unpacking icons...");
            Directory.CreateDirectory(CacheDirectory);
            var saved = 0;
            var inPack = new HashSet<ushort>();
            using var zip = ZipFile.OpenRead(tempPath);
            foreach (var entry in zip.Entries)
            {
                ct.ThrowIfCancellationRequested();
                // A zip made in the Mac Finder carries a __MACOSX/ copy ("._123.png") of every
                // file — resource-fork junk, not an icon.
                if (entry.FullName.StartsWith("__MACOSX/", StringComparison.Ordinal)
                    || !entry.Name.EndsWith(".png", StringComparison.OrdinalIgnoreCase)
                    || !ushort.TryParse(Path.GetFileNameWithoutExtension(entry.Name), out var iconId))
                {
                    continue;
                }

                entry.ExtractToFile(CachePath(iconId), overwrite: true);
                lock (SyncRoot)
                {
                    // Drop any "no icon" remembered from before, so it's read from disk next time.
                    if (MemoryCache.TryGetValue(iconId, out var cached) && cached is null)
                    {
                        MemoryCache.Remove(iconId);
                    }
                }

                inPack.Add(iconId);
                saved++;
            }

            ForgetIconsNotIn(inPack);
            return saved;
        }
        finally
        {
            try
            {
                File.Delete(tempPath);
            }
            catch (IOException)
            {
                // A leftover temp file is harmless.
            }
        }
    }

    /// <summary>
    /// The pack is the whole archive, so an icon it no longer has was taken off the site on
    /// purpose (hlwiki.com's scripts/removed-ik0ns.txt) — delete the copy saved from an older pack
    /// and drop it from the saved catalog, so it's gone here too. Skipped for a suspiciously small
    /// pack, which is more likely a broken download than thousands of removals.
    /// </summary>
    private static void ForgetIconsNotIn(HashSet<ushort> inPack)
    {
        if (inPack.Count < 1000)
        {
            return;
        }

        try
        {
            foreach (var path in Directory.EnumerateFiles(CacheDirectory, "*.png"))
            {
                if (ushort.TryParse(Path.GetFileNameWithoutExtension(path), out var id) && !inPack.Contains(id))
                {
                    File.Delete(path);
                    lock (SyncRoot)
                    {
                        MemoryCache.Remove(id);
                    }
                }
            }

            File.WriteAllLines(CatalogPath, inPack.Order().Select(id => id.ToString()));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Best-effort, same as the rest of the cache.
        }
    }

    private static string CatalogPath => Path.Combine(CacheDirectory, "catalog.txt");

    /// <summary>Whether this icon is already on disk — lets "Download all" skip what it has.</summary>
    public static bool IsCached(ushort iconId) => File.Exists(CachePath(iconId));

    /// <summary>
    /// Every icon number hlwiki.com has art for: the rows of its ik0ns.csv (see HotlineIconIndex),
    /// which lists the whole archive minus anything taken down. The numbers aren't contiguous
    /// (thousands of gaps), so the picker can't just count upward. Saved to disk each time, so
    /// with the site unreachable the picker still has the last list it saw.
    /// </summary>
    public static IReadOnlyList<ushort> CatalogFrom(IReadOnlyDictionary<ushort, HotlineIconInfo> index)
    {
        if (index.Count == 0)
        {
            return TryLoadCatalogFromDisk() ?? [];
        }

        var ids = index.Keys.Order().ToList();
        try
        {
            Directory.CreateDirectory(CacheDirectory);
            File.WriteAllLines(CatalogPath, ids.Select(id => id.ToString()));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Best-effort, same as the icon cache.
        }

        return ids;
    }

    private static List<ushort>? TryLoadCatalogFromDisk()
    {
        try
        {
            return File.Exists(CatalogPath)
                ? [.. File.ReadAllLines(CatalogPath).Select(l => ushort.TryParse(l, out var n) ? (int)n : -1).Where(n => n >= 0).Select(n => (ushort)n)]
                : null;
        }
        catch (IOException)
        {
            return null;
        }
    }
}
