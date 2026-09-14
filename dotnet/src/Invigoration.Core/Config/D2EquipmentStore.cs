using System.Text.Json;
using Invigoration.Core.Networking;
using Invigoration.Core.StatString;

namespace Invigoration.Core.Config;

/// <summary>What happened when fetching the Diablo II equipment map.</summary>
public enum D2EquipmentDownloadResult
{
    Saved,

    /// <summary>The server answered but doesn't serve the file.</summary>
    NotOnServer,

    /// <summary>The server sent something that isn't a map this Invigoration can read (wrong format, or a newer version).</summary>
    Unreadable,

    /// <summary>Couldn't reach the server, or the transfer broke off.</summary>
    Failed,
}

/// <summary>
/// The Diablo II equipment map (<see cref="D2EquipmentMap"/>) Invigoration keeps once it's
/// downloaded: fetched from a Command Center server over BNFTP only after the user says yes (the
/// user's own servers first, then us.bnet.cc — see <see cref="DownloadSources"/>), stored
/// in %AppData%/Invigoration/D2, and used on every server from then on — D2's item art never
/// changes, so there are no update checks. It also remembers a "no", so the offer isn't repeated.
/// </summary>
public static class D2EquipmentStore
{
    /// <summary>Where the download comes from by default, whichever server a bot is on.</summary>
    public const string DefaultHost = "us.bnet.cc";

    private static readonly Lock SyncRoot = new();
    private static bool _loaded;
    private static D2EquipmentMap? _current;
    private static bool? _declined;

    /// <summary>Raised after a new map is saved, so open views can re-describe characters.</summary>
    public static event Action? Changed;

    /// <summary>Test hook: points the store at a scratch folder.</summary>
    public static string? DirectoryOverride { get; set; }

    public static string Directory => DirectoryOverride ?? Path.Combine(ConfigStore.DefaultConfigDirectory(), "D2");

    public static string FilePath => Path.Combine(Directory, D2EquipmentMap.FileName);

    private static string ChoicePath => Path.Combine(Directory, "d2-equipment-choice.json");

    /// <summary>The stored map, or null if none has been downloaded (or the stored copy no longer reads).</summary>
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

    /// <summary>Checks and stores a downloaded map. Nothing is written unless it parses.</summary>
    /// <exception cref="FormatException">Not a map this version can read.</exception>
    public static void Save(byte[] json)
    {
        var map = D2EquipmentMap.Parse(json);
        lock (SyncRoot)
        {
            System.IO.Directory.CreateDirectory(Directory);
            var partial = FilePath + ".partial";
            File.WriteAllBytes(partial, json);
            File.Move(partial, FilePath, overwrite: true);
            _current = map;
            _loaded = true;
        }

        Changed?.Invoke();
    }

    /// <summary>
    /// Where to look for the map, in order: every non-Blizzard server these bots connect to (a
    /// Command Center node builds the file itself and serves it next to icons.bni), then
    /// <see cref="DefaultHost"/>. Official Battle.net never has it, so it isn't asked.
    /// </summary>
    public static IReadOnlyList<(string Host, int Port)> DownloadSources(IEnumerable<BotConfig> bots)
    {
        var sources = bots
            .Where(b => !string.IsNullOrWhiteSpace(b.BattlenetServer) &&
                        !b.BattlenetServer.Trim().EndsWith(".battle.net", StringComparison.OrdinalIgnoreCase))
            .Select(b => (Host: b.BattlenetServer.Trim(), Port: b.BattlenetPort > 0 ? b.BattlenetPort : BnftpClient.DefaultPort))
            .Distinct()
            .ToList();
        if (!sources.Any(s => s.Host.Equals(DefaultHost, StringComparison.OrdinalIgnoreCase)))
        {
            sources.Add((DefaultHost, BnftpClient.DefaultPort));
        }

        return sources;
    }

    /// <summary>Tries each source in turn and keeps the first map that downloads. Stops at an unreadable one too — a newer file format won't read any better from the next server. Never throws.</summary>
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

    /// <summary>Fetches the map from one Command Center server over BNFTP and stores it. Never throws: the outcome says what happened.</summary>
    public static async Task<(D2EquipmentDownloadResult Result, string Detail)> DownloadAsync(string host = DefaultHost, int port = BnftpClient.DefaultPort, TimeSpan? timeout = null)
    {
        BnftpFile? file;
        try
        {
            file = await BnftpClient.DownloadAsync(host, port, D2EquipmentMap.FileName, timeout ?? TimeSpan.FromSeconds(20)).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is IOException or System.Net.Sockets.SocketException or OperationCanceledException)
        {
            return (D2EquipmentDownloadResult.Failed, ex is OperationCanceledException ? "timed out" : ex.Message);
        }

        if (file is null)
        {
            return (D2EquipmentDownloadResult.NotOnServer, $"{host} doesn't serve {D2EquipmentMap.FileName}");
        }

        try
        {
            Save(file.Data);
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

    /// <summary>Test hook: forget what's loaded so the next read goes back to disk.</summary>
    public static void ResetCacheForTests()
    {
        lock (SyncRoot)
        {
            _loaded = false;
            _current = null;
            _declined = null;
        }
    }
}
