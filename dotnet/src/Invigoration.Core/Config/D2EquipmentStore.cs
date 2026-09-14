using System.IO.Compression;
using System.Text.Json;
using Invigoration.Core.Networking;
using Invigoration.Core.StatString;

namespace Invigoration.Core.Config;

/// <summary>What happened when fetching Diablo II data.</summary>
public enum D2EquipmentDownloadResult
{
    Saved,

    /// <summary>The server answered but doesn't serve the file.</summary>
    NotOnServer,

    /// <summary>The server sent something this Invigoration can't use (wrong format, or a newer version).</summary>
    Unreadable,

    /// <summary>Couldn't reach the server, or the transfer broke off.</summary>
    Failed,
}

/// <summary>
/// The Diablo II data Invigoration keeps once the user opts in, fetched over BNFTP from a trusted
/// Command Center server (<see cref="TrustedServers"/> — only us.bnet.cc for now) and stored in
/// %AppData%/Invigoration/D2 for use on every server:
/// </summary>
/// <remarks>
/// <para><c>d2-equipment.json</c> (<see cref="D2EquipmentMap"/>) — what the equipment bytes of a
/// character's statstring mean. Required: it's what opting in fetches first.</para>
/// <para><c>d2-characters.zip</c> — the character art pack, fetched alongside when the server has
/// it. Stored as-is and checked to be a readable zip; the App draws characters from it (D2CharacterLoader, via StatString.D2CharacterPack).</para>
/// <para>Each file's server file time is kept, so a bot connected to a trusted server can ask
/// SID_GETFILETIME and fetch only what's newer (BotEngine.D2Data.cs) — including the art pack
/// arriving for the first time on a server that didn't have it when the user opted in. A "no" to
/// the offer is remembered so it isn't repeated.</para>
/// </remarks>
public static class D2EquipmentStore
{
    /// <summary>Where the download comes from by default, whichever server a bot is on.</summary>
    public const string DefaultHost = "us.bnet.cc";

    public const string EquipmentFileName = D2EquipmentMap.FileName;

    public const string CharacterPackFileName = "d2-characters.zip";

    /// <summary>Every file kept here, in the order they're fetched.</summary>
    public static IReadOnlyList<string> TrackedFiles { get; } = [EquipmentFileName, CharacterPackFileName];

    private static readonly Lock SyncRoot = new();
    private static readonly HashSet<string> Refreshing = new(StringComparer.OrdinalIgnoreCase);
    private static bool _loaded;
    private static D2EquipmentMap? _current;
    private static bool? _declined;
    private static Dictionary<string, long>? _fileTimes;

    /// <summary>Raised after a file is saved, so open views can re-describe characters.</summary>
    public static event Action? Changed;

    /// <summary>Test hook: points the store at a scratch folder.</summary>
    public static string? DirectoryOverride { get; set; }

    public static string Directory => DirectoryOverride ?? Path.Combine(ConfigStore.DefaultConfigDirectory(), "D2");

    public static string FilePath => Path.Combine(Directory, EquipmentFileName);

    public static string CharacterPackPath => Path.Combine(Directory, CharacterPackFileName);

    private static string ChoicePath => Path.Combine(Directory, "d2-equipment-choice.json");

    private static string FileTimesPath => Path.Combine(Directory, "d2-file-times.json");

    /// <summary>The stored equipment map, or null if none has been downloaded (or the stored copy no longer reads).</summary>
    public static D2EquipmentMap? Current
    {
        get
        {
            lock (SyncRoot)
            {
                if (!_loaded)
                {
                    _loaded = true;
                    try
                    {
                        _current = File.Exists(FilePath) ? D2EquipmentMap.Parse(File.ReadAllBytes(FilePath)) : null;
                    }
                    catch (Exception ex) when (ex is IOException or FormatException or UnauthorizedAccessException)
                    {
                        _current = null;
                    }
                }

                return _current;
            }
        }
    }

    /// <summary>Whether the character art pack has been downloaded.</summary>
    public static bool HasCharacterPack => File.Exists(CharacterPackPath);

    /// <summary>Whether the user has opted in — anything is stored — which is what allows bots to check servers for newer copies.</summary>
    public static bool OptedIn => Current is not null || HasCharacterPack;

    /// <summary>Whether the user turned the download down. Only the automatic offer respects this; asking for it explicitly always works.</summary>
    public static bool Declined
    {
        get
        {
            lock (SyncRoot)
            {
                if (_declined is null)
                {
                    try
                    {
                        _declined = File.Exists(ChoicePath) &&
                                    JsonDocument.Parse(File.ReadAllText(ChoicePath)).RootElement.TryGetProperty("declined", out var d) &&
                                    d.GetBoolean();
                    }
                    catch (Exception ex) when (ex is IOException or JsonException or InvalidOperationException)
                    {
                        _declined = false;
                    }
                }

                return _declined.Value;
            }
        }

        set
        {
            lock (SyncRoot)
            {
                System.IO.Directory.CreateDirectory(Directory);
                File.WriteAllText(ChoicePath, JsonSerializer.Serialize(new { declined = value }));
                _declined = value;
            }
        }
    }

    /// <summary>The server file time recorded when <paramref name="fileName"/> was last downloaded, or 0 if it never was.</summary>
    public static long StoredFileTime(string fileName)
    {
        lock (SyncRoot)
        {
            return LoadFileTimes().GetValueOrDefault(fileName);
        }
    }

    /// <summary>
    /// Whether a server's SID_GETFILETIME answer means there's something to fetch: the user has
    /// opted in, it's a file kept here, the server has it (time 0 means it doesn't), and its copy
    /// is newer than ours — or ours is missing, like an art pack the first server didn't have.
    /// </summary>
    public static bool IsNewerOnServer(string fileName, long serverFileTime) =>
        OptedIn &&
        TrackedFiles.Contains(fileName, StringComparer.OrdinalIgnoreCase) &&
        serverFileTime > 0 &&
        (serverFileTime > StoredFileTime(fileName) || !File.Exists(Path.Combine(Directory, fileName)));

    /// <summary>Checks and stores a downloaded equipment map. Nothing is written unless it parses.</summary>
    /// <exception cref="FormatException">Not a map this version can read.</exception>
    public static void Save(byte[] json, long fileTime = 0)
    {
        var map = D2EquipmentMap.Parse(json);
        lock (SyncRoot)
        {
            WriteFile(EquipmentFileName, json, fileTime);
            _current = map;
            _loaded = true;
        }

        Changed?.Invoke();
    }

    /// <summary>Checks and stores a downloaded character art pack: it has to open as a zip with at least one entry.</summary>
    /// <exception cref="FormatException">Not a readable zip.</exception>
    public static void SaveCharacterPack(byte[] zip, long fileTime = 0)
    {
        try
        {
            using var archive = new ZipArchive(new MemoryStream(zip), ZipArchiveMode.Read);
            if (archive.Entries.Count == 0)
            {
                throw new FormatException("The character pack is an empty zip.");
            }
        }
        catch (InvalidDataException ex)
        {
            throw new FormatException("The character pack isn't a readable zip.", ex);
        }

        lock (SyncRoot)
        {
            WriteFile(CharacterPackFileName, zip, fileTime);
        }

        Changed?.Invoke();
    }

    /// <summary>
    /// The only servers this feature ever talks to — for downloads and for SID_GETFILETIME update
    /// checks alike. Deliberately just us.bnet.cc for now: other servers (PvPGN, Atlas, anything
    /// else) never get either request, so an implementation that mishandles an unfamiliar file
    /// request can't be knocked over by it. Other hosts can be added once the file and its
    /// documentation are published for other servers to serve.
    /// </summary>
    public static IReadOnlyList<(string Host, int Port)> TrustedServers { get; } = [(DefaultHost, BnftpClient.DefaultPort)];

    /// <summary>Whether a bot's server is one of <see cref="TrustedServers"/>.</summary>
    public static bool IsTrustedServer(string server) =>
        TrustedServers.Any(s => s.Host.Equals(server.Trim(), StringComparison.OrdinalIgnoreCase));

    /// <summary>Tries each source in turn and keeps the first that has the equipment map. Stops at an unreadable one too — a newer file format won't read any better from the next server. Never throws.</summary>
    public static async Task<(D2EquipmentDownloadResult Result, string Detail)> DownloadAsync(IReadOnlyList<(string Host, int Port)> sources, TimeSpan? timeout = null)
    {
        var misses = new List<string>();
        var result = D2EquipmentDownloadResult.NotOnServer;
        foreach (var (host, port) in sources)
        {
            var (attempt, detail) = await DownloadAsync(host, port, timeout).ConfigureAwait(false);
            if (attempt is D2EquipmentDownloadResult.Saved or D2EquipmentDownloadResult.Unreadable)
            {
                return (attempt, detail);
            }

            // "Not there" is only the overall answer if every source said so; any failure to connect wins.
            if (attempt == D2EquipmentDownloadResult.Failed)
            {
                result = D2EquipmentDownloadResult.Failed;
            }

            misses.Add($"{host}: {detail}");
        }

        return (result, misses.Count == 0 ? "no servers to ask" : string.Join("; ", misses));
    }

    /// <summary>Fetches the equipment map from one server, then the character art pack if that server has it. The result is the map's; the pack is extra. Never throws.</summary>
    public static async Task<(D2EquipmentDownloadResult Result, string Detail)> DownloadAsync(string host = DefaultHost, int port = BnftpClient.DefaultPort, TimeSpan? timeout = null)
    {
        var (result, detail) = await FetchAsync(host, port, EquipmentFileName, timeout).ConfigureAwait(false);
        if (result != D2EquipmentDownloadResult.Saved)
        {
            return (result, detail);
        }

        var (packResult, packDetail) = await FetchAsync(host, port, CharacterPackFileName, timeout).ConfigureAwait(false);
        return packResult == D2EquipmentDownloadResult.Saved
            ? (result, $"{detail}; character art {packDetail}")
            : (result, detail);
    }

    /// <summary>
    /// Re-downloads one kept file from the server that reported a newer copy. A second bot noticing
    /// the same update while the first is still fetching it is ignored rather than fetching twice.
    /// </summary>
    public static async Task<(D2EquipmentDownloadResult Result, string Detail)> RefreshFileAsync(string host, int port, string fileName, TimeSpan? timeout = null)
    {
        lock (SyncRoot)
        {
            if (!Refreshing.Add(fileName))
            {
                return (D2EquipmentDownloadResult.Failed, "already being updated");
            }
        }

        try
        {
            return await FetchAsync(host, port, fileName, timeout).ConfigureAwait(false);
        }
        finally
        {
            lock (SyncRoot)
            {
                Refreshing.Remove(fileName);
            }
        }
    }

    private static async Task<(D2EquipmentDownloadResult Result, string Detail)> FetchAsync(string host, int port, string fileName, TimeSpan? timeout)
    {
        BnftpFile? file;
        try
        {
            file = await BnftpClient.DownloadAsync(host, port, fileName, timeout ?? TimeSpan.FromSeconds(60)).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is IOException or System.Net.Sockets.SocketException or OperationCanceledException)
        {
            return (D2EquipmentDownloadResult.Failed, ex is OperationCanceledException ? "timed out" : ex.Message);
        }

        if (file is null)
        {
            return (D2EquipmentDownloadResult.NotOnServer, $"{host} doesn't serve {fileName}");
        }

        try
        {
            if (fileName.Equals(CharacterPackFileName, StringComparison.OrdinalIgnoreCase))
            {
                SaveCharacterPack(file.Data, file.FileTime);
            }
            else
            {
                Save(file.Data, file.FileTime);
            }

            return (D2EquipmentDownloadResult.Saved, $"{file.Data.Length:N0} bytes from {host}");
        }
        catch (FormatException ex)
        {
            return (D2EquipmentDownloadResult.Unreadable, ex.Message);
        }
        catch (IOException ex)
        {
            return (D2EquipmentDownloadResult.Failed, ex.Message);
        }
    }

    /// <summary>Atomic write, and records the server's file time for later SID_GETFILETIME comparisons. Caller holds SyncRoot.</summary>
    private static void WriteFile(string fileName, byte[] bytes, long fileTime)
    {
        System.IO.Directory.CreateDirectory(Directory);
        var path = Path.Combine(Directory, fileName);
        var partial = path + ".partial";
        File.WriteAllBytes(partial, bytes);
        File.Move(partial, path, overwrite: true);

        var times = LoadFileTimes();
        times[fileName] = fileTime;
        File.WriteAllText(FileTimesPath, JsonSerializer.Serialize(times));
    }

    private static Dictionary<string, long> LoadFileTimes()
    {
        if (_fileTimes is not null)
        {
            return _fileTimes;
        }

        try
        {
            _fileTimes = File.Exists(FileTimesPath)
                ? JsonSerializer.Deserialize<Dictionary<string, long>>(File.ReadAllText(FileTimesPath)) ?? []
                : [];
        }
        catch (Exception ex) when (ex is IOException or JsonException)
        {
            _fileTimes = [];
        }

        _fileTimes = new Dictionary<string, long>(_fileTimes, StringComparer.OrdinalIgnoreCase);
        return _fileTimes;
    }

    /// <summary>Test hook: forget what's loaded so the next read goes back to disk.</summary>
    public static void ResetCacheForTests()
    {
        lock (SyncRoot)
        {
            _loaded = false;
            _current = null;
            _declined = null;
            _fileTimes = null;
            Refreshing.Clear();
        }
    }
}
