using System.Collections.Concurrent;
using Invigoration.Core.Config;

namespace Invigoration.Core.Sc2;

/// <summary>
/// StarCraft II profile portraits: 17 sheets of 6×6 portraits (152px each), Blizzard's art as
/// ncarrillo/superiority (MIT) packages them. Not shipped with Invigoration: the user is offered a
/// one-time download (about 21 MB) the first time a StarCraft II bot connects, from superiority's
/// GitHub at a pinned commit, into the settings folder.
/// </summary>
public static class Sc2PortraitStore
{
    public const int SheetCount = 17;
    public const int CellSize = 152;
    public const int SheetColumns = 6;

    /// <summary>The superiority commit the sheets are downloaded from, so a later change there can't break them.</summary>
    private const string Commit = "d312643305790a4215171f6717741a8e26e9e8c9";

    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(2) };

    public static string Folder => Path.Combine(ConfigStore.DefaultConfigDirectory(), "Sc2Portraits");

    /// <summary>Raised once the sheets are on disk, so user lists can redraw.</summary>
    public static event Action? Downloaded;

    public static string SheetPath(int sheet) => Path.Combine(Folder, $"atlas-{sheet:00}.png");

    public static bool IsDownloaded =>
        Enumerable.Range(0, SheetCount).All(i => new FileInfo(SheetPath(i)) is { Exists: true, Length: > 0 });

    /// <summary>The user said not to ask again.</summary>
    public static bool Declined => File.Exists(Path.Combine(Folder, "declined"));

    public static void Decline()
    {
        Directory.CreateDirectory(Folder);
        File.WriteAllText(Path.Combine(Folder, "declined"), "The user chose not to download StarCraft II portraits. Delete this file to be asked again.");
    }

    /// <summary>Downloads any missing sheet. Each goes to a .part file first, so an interrupted download is never mistaken for a sheet.</summary>
    public static async Task DownloadAsync(IProgress<int>? sheetsDone, CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(Folder);
        for (var sheet = 0; sheet < SheetCount; sheet++)
        {
            var path = SheetPath(sheet);
            if (new FileInfo(path) is not { Exists: true, Length: > 0 })
            {
                var url = $"https://raw.githubusercontent.com/ncarrillo/superiority/{Commit}/app/macos/resources/images/portrait-atlases/atlas-{sheet:00}.png";
                var bytes = await Http.GetByteArrayAsync(url, cancellationToken).ConfigureAwait(false);
                await File.WriteAllBytesAsync(path + ".part", bytes, cancellationToken).ConfigureAwait(false);
                File.Move(path + ".part", path, overwrite: true);
            }

            sheetsDone?.Report(sheet + 1);
        }

        Downloaded?.Invoke();
    }
}

/// <summary>
/// Which portrait each StarCraft II chat member (and friend, by name) has, as a sheet and a cell,
/// learned by the native SC2 client. Like NativeMemberProducts, keyed by name across bots. Kept on
/// disk too, so on reconnect everyone seen before shows at once; a later lookup corrects anyone
/// who has changed theirs.
/// </summary>
public static class NativeMemberPortraits
{
    private static readonly ConcurrentDictionary<string, (ushort Sheet, ushort Cell)> ByName = new(StringComparer.OrdinalIgnoreCase);
    private static readonly ConcurrentDictionary<string, string> Details = new(StringComparer.OrdinalIgnoreCase);
    private static readonly ConcurrentDictionary<string, Stimpak.Presence> States = new(StringComparer.OrdinalIgnoreCase);
    private static readonly Timer SaveTimer = new(_ => Save());
    private static int _loaded;

    /// <summary>Raised when someone's portrait or detail becomes known or changes, so user lists redraw.</summary>
    public static event Action? Changed;

    private static string FilePath => Path.Combine(Sc2PortraitStore.Folder, "known-portraits.json");

    public static void Set(string name, ushort sheet, ushort cell)
    {
        EnsureLoaded();
        if (name.Length == 0 || sheet >= Sc2PortraitStore.SheetCount || cell >= Sc2PortraitStore.SheetColumns * Sc2PortraitStore.SheetColumns)
        {
            return;
        }

        if (!ByName.TryGetValue(name, out var known) || known != (sheet, cell))
        {
            ByName[name] = (sheet, cell);
            SaveTimer.Change(TimeSpan.FromSeconds(10), Timeout.InfiniteTimeSpan);
            Changed?.Invoke();
        }
    }

    public static (ushort Sheet, ushort Cell)? For(string name)
    {
        EnsureLoaded();
        return ByName.TryGetValue(name, out var portrait) ? portrait : null;
    }

    /// <summary>A short line about a member for the user list's Full view: their BattleTag, and whether they're in a game, away or busy.</summary>
    public static void SetDetail(string name, string detail)
    {
        if (name.Length > 0 && (!Details.TryGetValue(name, out var known) || known != detail))
        {
            Details[name] = detail;
            Changed?.Invoke();
        }
    }

    public static string DetailFor(string name) => Details.TryGetValue(name, out var detail) ? detail : "";

    /// <summary>A member's status as their presence says (away, busy, in a game), for the user list's dot.</summary>
    public static void SetState(string name, Stimpak.Presence state)
    {
        if (name.Length > 0 && (!States.TryGetValue(name, out var known) || known != state))
        {
            States[name] = state;
            Changed?.Invoke();
        }
    }

    public static Stimpak.Presence? StateFor(string name) => States.TryGetValue(name, out var state) ? state : null;

    private static void EnsureLoaded()
    {
        if (Interlocked.Exchange(ref _loaded, 1) == 1)
        {
            return;
        }

        try
        {
            if (File.Exists(FilePath)
                && System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, ushort[]>>(File.ReadAllText(FilePath)) is { } saved)
            {
                foreach (var (name, value) in saved)
                {
                    if (value is [var sheet, var cell])
                    {
                        ByName.TryAdd(name, (sheet, cell));
                    }
                }
            }
        }
        catch (Exception ex) when (ex is IOException or System.Text.Json.JsonException or UnauthorizedAccessException)
        {
            // Only a head start; lookups fill it in again.
        }
    }

    private static void Save()
    {
        try
        {
            Directory.CreateDirectory(Sc2PortraitStore.Folder);
            var snapshot = ByName.ToDictionary(p => p.Key, p => new[] { p.Value.Sheet, p.Value.Cell });
            File.WriteAllText(FilePath, System.Text.Json.JsonSerializer.Serialize(snapshot));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Tried again on the next change.
        }
    }
}
