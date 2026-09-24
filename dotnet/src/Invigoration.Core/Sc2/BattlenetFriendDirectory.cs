using System.Collections.Concurrent;
using System.Text.Json;
using Invigoration.Core.Config;

namespace Invigoration.Core.Sc2;

/// <summary>
/// The Battle.net friends list as StarCraft: Remastered receives it (BattleTag, real name, the game
/// they're in), shared per Battle.net profile with the SC2 bot, which Battle.net sends no BattleTags.
/// Names are kept on disk next to the profile's saved sign-ins so SC2 has them without SC:R
/// connected; what friends are doing is only known while an SC:R bot on the profile is.
/// </summary>
public static class BattlenetFriendDirectory
{
    public sealed record Friend(string BattleTag, string RealName);

    public sealed record Activity(bool Online, string Program, string Detail);

    private static readonly ConcurrentDictionary<string, Dictionary<uint, Friend>> Names = new();
    private static readonly ConcurrentDictionary<string, Dictionary<uint, Activity>> Live = new();

    /// <summary>Raised with the profile ID when its friends change.</summary>
    public static event Action<string>? Changed;

    public static void Update(string profileId, IEnumerable<(uint AccountId, string BattleTag, string RealName, Activity Activity)> friends)
    {
        var names = new Dictionary<uint, Friend>();
        var live = new Dictionary<uint, Activity>();
        foreach (var (id, tag, realName, activity) in friends)
        {
            names[id] = new Friend(tag, realName);
            live[id] = activity;
        }

        Live[profileId] = live;
        var before = Load(profileId);
        Names[profileId] = names;
        if (!before.OrderBy(p => p.Key).SequenceEqual(names.OrderBy(p => p.Key)))
        {
            Save(profileId, names);
        }

        Changed?.Invoke(profileId);
    }

    /// <summary>SC:R disconnected: what friends are doing is no longer known.</summary>
    public static void ForgetActivity(string profileId)
    {
        if (Live.TryRemove(profileId, out _))
        {
            Changed?.Invoke(profileId);
        }
    }

    public static Friend? Find(string profileId, uint accountId) =>
        Load(profileId).TryGetValue(accountId, out var friend) ? friend : null;

    public static Activity? ActivityOf(string profileId, uint accountId) =>
        Live.TryGetValue(profileId, out var live) && live.TryGetValue(accountId, out var activity) ? activity : null;

    private static Dictionary<uint, Friend> Load(string profileId) =>
        Names.GetOrAdd(profileId, id =>
        {
            try
            {
                var path = FilePath(id);
                return File.Exists(path)
                    ? JsonSerializer.Deserialize<Dictionary<uint, Friend>>(File.ReadAllText(path)) ?? []
                    : [];
            }
            catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
            {
                return [];
            }
        });

    private static void Save(string profileId, Dictionary<uint, Friend> names)
    {
        try
        {
            var path = FilePath(profileId);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, JsonSerializer.Serialize(names));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Only a convenience for SC2; it's rebuilt next time SC:R signs in.
        }
    }

    private static string FilePath(string profileId) =>
        Path.ChangeExtension(BattlenetCredentialProfileStore.CredentialFilePath(profileId), ".friends.json");
}
