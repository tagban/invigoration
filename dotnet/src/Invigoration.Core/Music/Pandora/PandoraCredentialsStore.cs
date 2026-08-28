using System.Text.Json;

namespace Invigoration.Core.Music.Pandora;

/// <summary>
/// The user's own Pandora account username/password — global, not per-bot, same reasoning as
/// MusicSettingsStore (one shared player for the whole app). Plaintext at rest: this codebase
/// already stores Battle.net credentials the same way in BotConfig, so a new, different-strength
/// scheme for this one store would be inconsistent without actually improving security (the file
/// itself, not the encoding, is what needs protecting).
/// </summary>
public static class PandoraCredentialsStore
{
    private static string FilePath => Path.Combine(Config.ConfigStore.DefaultConfigDirectory(), "pandora-credentials.json");

    private static StoredCredentials? _cached;

    public static string Username
    {
        get => Current.Username;
        set => Save(Current with { Username = value });
    }

    public static string Password
    {
        get => Current.Password;
        set => Save(Current with { Password = value });
    }

    public static bool HasCredentials => Current.Username.Length > 0 && Current.Password.Length > 0;

    private static StoredCredentials Current
    {
        get
        {
            if (_cached is { } cached)
            {
                return cached;
            }

            var loaded = Load();
            _cached = loaded;
            return loaded;
        }
    }

    private static void Save(StoredCredentials credentials)
    {
        _cached = credentials;
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(credentials));
        }
        catch (IOException)
        {
            // Best-effort, same as MusicSettingsStore — the in-memory cache still reflects the
            // change for the rest of this run even if the write itself failed.
        }
    }

    private static StoredCredentials Load()
    {
        try
        {
            if (!File.Exists(FilePath))
            {
                return new StoredCredentials("", "");
            }

            var stored = JsonSerializer.Deserialize<StoredCredentials>(File.ReadAllText(FilePath));
            return stored ?? new StoredCredentials("", "");
        }
        catch (Exception ex) when (ex is IOException or JsonException)
        {
            return new StoredCredentials("", "");
        }
    }

    private sealed record StoredCredentials(string Username, string Password);
}
