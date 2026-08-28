using System.Text.Json;

namespace Invigoration.Core.Hotline;

/// <summary>The list of top-level Hotline trackers the user has added — persisted at %AppData%/Invigoration/hotline-trackers.json, same shape as ConfigStore (bots.json) but for this separate, non-Battle.net entity type.</summary>
public sealed class HotlineTrackerConfigStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public string FilePath { get; }

    public HotlineTrackerConfigStore(string? filePath = null)
    {
        FilePath = filePath ?? Path.Combine(Config.ConfigStore.DefaultConfigDirectory(), "hotline-trackers.json");
    }

    public List<HotlineTrackerConfig> Load()
    {
        if (!File.Exists(FilePath))
        {
            return [];
        }

        var json = File.ReadAllText(FilePath);
        return JsonSerializer.Deserialize<List<HotlineTrackerConfig>>(json, JsonOptions) ?? [];
    }

    public void Save(IReadOnlyList<HotlineTrackerConfig> trackers)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        File.WriteAllText(FilePath, JsonSerializer.Serialize(trackers, JsonOptions));
    }
}
