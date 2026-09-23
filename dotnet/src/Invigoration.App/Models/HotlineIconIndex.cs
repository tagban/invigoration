namespace Invigoration.App.Models;

/// <summary>What one icon shows, for searching: its lettering, what's pictured, and its main colors.</summary>
public sealed record HotlineIconInfo(string Text, IReadOnlyList<string> Tags, IReadOnlyList<string> Colors)
{
    /// <summary>Every word in the lettering, tags and colors, lowercased — a search word matches when one of these starts with it.</summary>
    public IReadOnlyList<string> Words { get; } =
        [.. Split(Text).Concat(Tags.SelectMany(Split)).Concat(Colors.SelectMany(Split)).Distinct()];

    /// <summary>Tagged "nsfw" in the index — hidden in the picker unless Show NSFW is ticked.</summary>
    public bool IsNsfw { get; } = Tags.Any(t => t.Equals("nsfw", StringComparison.OrdinalIgnoreCase));

    /// <summary>Lettering, then tags, then colors — the tooltip on a picker cell.</summary>
    public string Describe()
    {
        var parts = new List<string>();
        if (Text.Length > 0)
        {
            parts.Add($"\"{Text}\"");
        }

        if (Tags.Count > 0)
        {
            parts.Add(string.Join(", ", Tags));
        }

        if (Colors.Count > 0)
        {
            parts.Add(string.Join(", ", Colors));
        }

        return string.Join("\n", parts);
    }

    public static IEnumerable<string> Split(string text) =>
        text.ToLowerInvariant().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
            .Select(w => w.Trim('"', '\'', '.', ',', '!', '?', ':', ';', '(', ')', '-', '*'))
            .Where(w => w.Length > 0);
}

/// <summary>
/// The search index hlwiki.com publishes beside the icons (<see cref="IndexUrl"/>), a plain CSV
/// the site's own gallery searches with too, and that anyone can open in a spreadsheet or fix.
/// Built once from the whole archive: the lettering on each icon, what it pictures (a flag, an
/// eye, a franchise logo...), and its main colors. Cached to disk like the catalog, so search
/// still works offline once it's been fetched.
/// </summary>
public static class HotlineIconIndex
{
    public const string IndexUrl = "https://hlwiki.com/ik0ns/ik0ns.csv";

    private static readonly HttpClient Http = new();

    private static string IndexPath => Path.Combine(HotlineIconLoader.CacheDirectory, "ik0ns.csv");

    /// <summary>Fetches the index (or reads the saved copy when the site can't be reached). Empty when neither exists — search then falls back to icon numbers alone.</summary>
    public static async Task<IReadOnlyDictionary<ushort, HotlineIconInfo>> LoadAsync(CancellationToken ct = default)
    {
        try
        {
            var csv = await Http.GetStringAsync(IndexUrl, ct).ConfigureAwait(false);
            var parsed = Parse(csv);
            if (parsed.Count > 0)
            {
                try
                {
                    Directory.CreateDirectory(HotlineIconLoader.CacheDirectory);
                    await File.WriteAllTextAsync(IndexPath, csv, ct).ConfigureAwait(false);
                }
                catch (IOException)
                {
                    // Best-effort, same as the icon cache.
                }

                return parsed;
            }
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
        {
            // Fall back to the saved copy.
        }

        try
        {
            return File.Exists(IndexPath) ? Parse(await File.ReadAllTextAsync(IndexPath, ct).ConfigureAwait(false)) : new Dictionary<ushort, HotlineIconInfo>();
        }
        catch (IOException)
        {
            return new Dictionary<ushort, HotlineIconInfo>();
        }
    }

    /// <summary>
    /// <c>id,text,tags,colors</c> with a header row; tags and colors are each separated by
    /// semicolons. Standard CSV quoting (a field in double quotes may hold commas, newlines and
    /// doubled quotes), since lettering can contain any of those.
    /// </summary>
    public static Dictionary<ushort, HotlineIconInfo> Parse(string csv)
    {
        var result = new Dictionary<ushort, HotlineIconInfo>();
        var first = true;
        foreach (var row in ReadRows(csv))
        {
            if (first)
            {
                first = false;
                if (row.Count > 0 && row[0].Equals("id", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
            }

            if (row.Count == 0 || !ushort.TryParse(row[0].Trim(), out var id))
            {
                continue;
            }

            result[id] = new HotlineIconInfo(
                Field(row, 1).Trim(),
                SplitList(Field(row, 2)),
                SplitList(Field(row, 3)));
        }

        return result;
    }

    private static string Field(List<string> row, int index) => index < row.Count ? row[index] : "";

    private static List<string> SplitList(string field) =>
        [.. field.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)];

    private static IEnumerable<List<string>> ReadRows(string csv)
    {
        var row = new List<string>();
        var field = new System.Text.StringBuilder();
        var quoted = false;
        for (var i = 0; i < csv.Length; i++)
        {
            var c = csv[i];
            if (quoted)
            {
                if (c == '"')
                {
                    if (i + 1 < csv.Length && csv[i + 1] == '"')
                    {
                        field.Append('"');
                        i++;
                    }
                    else
                    {
                        quoted = false;
                    }
                }
                else
                {
                    field.Append(c);
                }
            }
            else if (c == '"')
            {
                quoted = true;
            }
            else if (c == ',')
            {
                row.Add(field.ToString());
                field.Clear();
            }
            else if (c is '\n' or '\r')
            {
                if (c == '\r' && i + 1 < csv.Length && csv[i + 1] == '\n')
                {
                    i++;
                }

                row.Add(field.ToString());
                field.Clear();
                yield return row;
                row = [];
            }
            else
            {
                field.Append(c);
            }
        }

        if (field.Length > 0 || row.Count > 0)
        {
            row.Add(field.ToString());
            yield return row;
        }
    }

    /// <summary>
    /// Whether an icon fits what's typed. All digits means an icon number (41 finds 41, 410-419,
    /// 4100...), as before. Otherwise every word has to start one of the icon's words, so "red
    /// flag" finds red flags, and "eye" finds "eyes" without "red" finding "tired".
    /// </summary>
    public static bool Matches(ushort id, HotlineIconInfo? info, IReadOnlyList<string> queryWords, string rawQuery)
    {
        if (rawQuery.All(char.IsAsciiDigit))
        {
            return id.ToString().StartsWith(rawQuery, StringComparison.Ordinal);
        }

        if (info is null)
        {
            return false;
        }

        return queryWords.All(q => info.Words.Any(w => w.StartsWith(q, StringComparison.Ordinal)));
    }
}
