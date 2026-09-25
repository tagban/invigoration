using Invigoration.Sc2.Protobuf;

namespace Invigoration.Scr;

/// <summary>One named stat of a character (ToonProfile.GetStats), e.g. legacy_wins = 12.</summary>
public sealed record ScrStat(string Name, ulong Value, uint Extra);

/// <summary>
/// ToonProfile.GetStats: a character's stats by name. Method IDs and both layouts were read from
/// the retail client's library; the values are the ones the game's own (only) GetStats call
/// passes, read by emulating that call: no program, the player's gateway, the name, ".*",
/// and 0xFFFFFFFF in field 5.
/// </summary>
public static class ScrStats
{
    public const uint Service = 0x8B952A18;
    public const uint GetStatsMethod = 0x81F5E3C9;

    /// <summary>
    /// GetStatsRequest: {1: program (a FourCC packed big-endian, sent only when given), 2: gateway
    /// (only when not 0), 3: character name (required), 4: stat-name regex (the client's default is
    /// ".*"), 5: uint32 of unknown meaning}.
    /// </summary>
    public static byte[] Request(string characterName, string? program, uint gateway, string filter = ".*", uint field5 = uint.MaxValue)
    {
        var request = new ProtoWriter();
        if (!string.IsNullOrEmpty(program))
        {
            request.WriteUInt32(1, Pack(program));
        }

        if (gateway != 0)
        {
            request.WriteUInt64(2, gateway);
        }

        request.WriteString(3, characterName);
        request.WriteString(4, filter);
        request.WriteUInt32(5, field5);
        return request.ToArray();
    }

    /// <summary>GetStatsResponse: {1: repeated Stat {1: name, 2: uint64 value, 3: uint32}, 2: uint32}.</summary>
    public static (IReadOnlyList<ScrStat> Stats, uint Status) Decode(byte[] body)
    {
        var stats = new List<ScrStat>();
        uint status = 0;
        var outer = new ProtoReader(body);
        while (outer.HasMore)
        {
            var (field, type) = outer.ReadTag();
            if (field == 1 && type == WireType.LengthDelimited)
            {
                string name = "";
                ulong value = 0;
                uint extra = 0;
                var r = new ProtoReader(outer.ReadLengthDelimited());
                while (r.HasMore)
                {
                    var (inner, innerType) = r.ReadTag();
                    switch (inner)
                    {
                        case 1 when innerType == WireType.LengthDelimited:
                            name = r.ReadString();
                            break;
                        case 2 when innerType == WireType.Varint:
                            value = r.ReadVarint();
                            break;
                        case 3 when innerType == WireType.Varint:
                            extra = (uint)r.ReadVarint();
                            break;
                        default:
                            r.Skip(innerType);
                            break;
                    }
                }

                stats.Add(new ScrStat(name, value, extra));
            }
            else if (field == 2 && type == WireType.Varint)
            {
                status = (uint)outer.ReadVarint();
            }
            else
            {
                outer.Skip(type);
            }
        }

        return (stats, status);
    }

    /// <summary>
    /// A line for chat: wins, losses, draws and disconnects as overall, ranked (mm_) and classic
    /// (legacy_) groups, then games, play time, APM and when the character was made. Other stats
    /// follow by name. Keys seen live: legacy_wins, legacy_losses, legacy_disconnects and
    /// legacy_toon_creation_time (a Windows FILETIME).
    /// </summary>
    public static string Describe(IReadOnlyList<ScrStat> stats)
    {
        var byName = stats.GroupBy(s => s.Name, StringComparer.OrdinalIgnoreCase).ToDictionary(g => g.Key, g => g.First().Value, StringComparer.OrdinalIgnoreCase);
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var parts = new List<string>();

        foreach (var (prefix, label) in new[] { ("", ""), ("mm_", "ranked"), ("legacy_", "classic") })
        {
            var record = new List<string>();
            foreach (var (key, word) in new[] { ("wins", "wins"), ("losses", "losses"), ("draws", "draws"), ("disconnects", "disconnects") })
            {
                if (byName.TryGetValue(prefix + key, out var value))
                {
                    used.Add(prefix + key);
                    if (value != 0 || key is "wins" or "losses")
                    {
                        record.Add($"{value:N0} {word}");
                    }
                }
            }

            if (record.Count > 0)
            {
                parts.Add(label.Length > 0 ? $"{label}: {string.Join(", ", record)}" : string.Join(", ", record));
            }
        }

        if (Take("games_played") is { } games)
        {
            parts.Add($"{games:N0} games");
        }

        if (Take("play_time") is { } seconds && seconds > 0)
        {
            parts.Add($"played {TimeSpan.FromSeconds(seconds).TotalHours:N0} h");
        }

        if (Take("APM") is { } apm && apm > 0)
        {
            parts.Add($"{apm} APM");
        }

        if (Take("legacy_toon_creation_time") is { } created && created > 0)
        {
            try
            {
                parts.Add($"created {DateTime.FromFileTimeUtc((long)created):d MMM yyyy}");
            }
            catch (ArgumentOutOfRangeException)
            {
            }
        }

        parts.AddRange(stats.Where(s => !used.Contains(s.Name) && s.Value != 0).Select(s => $"{s.Name} {s.Value:N0}"));
        return parts.Count > 0 ? string.Join(" · ", parts) : "no games recorded";

        ulong? Take(string key)
        {
            used.Add(key);
            return byName.TryGetValue(key, out var value) ? value : null;
        }
    }

    /// <summary>"S1" → 0x5331, "SEXP" → 0x53455850: big-endian, as the client packs it.</summary>
    public static uint Pack(string program) => program.Aggregate(0u, (packed, c) => (packed << 8) | (byte)c);
}
