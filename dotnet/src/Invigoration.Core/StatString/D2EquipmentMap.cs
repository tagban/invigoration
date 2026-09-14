using System.Text.Json;

namespace Invigoration.Core.StatString;

/// <summary>One thing a Diablo II character is wearing: which slot, what it looks like (every item that shares that look), and its tint if it has one.</summary>
public sealed record D2WornItem(string Slot, IReadOnlyList<string> Looks, string? Tint)
{
    /// <summary>"Cap / War Hat / Shako (Crystal Blue)" — at most three names, since several items can share one look.</summary>
    public override string ToString()
    {
        var names = Looks.Count <= 3 ? string.Join(" / ", Looks) : string.Join(" / ", Looks.Take(3)) + " / …";
        return Tint is null ? names : $"{names} ({Tint})";
    }
}

/// <summary>
/// What the equipment bytes of a Diablo II character's chat statstring mean — Command Center's
/// <c>d2-equipment.json</c> (bnet_command_center docs/D2-EQUIPMENT-FILE.md), which its server
/// builds from the game install and serves over BNFTP next to icons.bni. The bytes aren't item ids:
/// they index a graphics table the game builds from its item tables, so a value names a look that
/// several items can share (a Cap, a War Hat and a Shako look the same).
/// </summary>
public sealed class D2EquipmentMap
{
    public const string FileName = "d2-equipment.json";
    public const string FormatName = "bnetcc-d2-equipment";
    public const int SupportedVersion = 1;

    private const byte None = 255;

    private static readonly string[] BodyArmorSlots = ["torso", "legs", "right_arm", "left_arm", "right_shoulder", "left_shoulder"];

    private readonly Dictionary<string, Slot> _slots;
    private readonly Dictionary<string, IReadOnlyList<string>> _bodyArmorSets;
    private readonly IReadOnlyList<string> _colors;

    private sealed record Slot(int Offset, int TintOffset, Dictionary<byte, IReadOnlyList<string>> Items, Dictionary<byte, string> Weights);

    private D2EquipmentMap(Dictionary<string, Slot> slots, Dictionary<string, IReadOnlyList<string>> bodyArmorSets, IReadOnlyList<string> colors)
    {
        _slots = slots;
        _bodyArmorSets = bodyArmorSets;
        _colors = colors;
    }

    /// <summary>Reads the file. It's data from the network, so everything is checked: the format name, the version, and that every offset lands inside the 33-byte portrait.</summary>
    /// <exception cref="FormatException">Not a d2-equipment.json this version understands.</exception>
    public static D2EquipmentMap Parse(ReadOnlySpan<byte> json)
    {
        try
        {
            using var doc = JsonDocument.Parse(json.ToArray());
            var root = doc.RootElement;
            if (!root.TryGetProperty("format", out var format) || format.GetString() != FormatName)
            {
                throw new FormatException("Not a Command Center Diablo II equipment file.");
            }

            if (!root.TryGetProperty("version", out var version) || version.GetInt32() != SupportedVersion)
            {
                throw new FormatException($"Equipment file version {(root.TryGetProperty("version", out var v) ? v.ToString() : "?")} isn't one this Invigoration understands (it reads version {SupportedVersion}).");
            }

            var slots = new Dictionary<string, Slot>(StringComparer.Ordinal);
            foreach (var slot in root.GetProperty("slots").EnumerateArray())
            {
                var offset = slot.GetProperty("offset").GetInt32();
                var tintOffset = slot.GetProperty("tint_offset").GetInt32();
                if (offset is < 0 or >= D2Character.PortraitLength || tintOffset is < 0 or >= D2Character.PortraitLength)
                {
                    throw new FormatException("Equipment slot offset outside the portrait.");
                }

                var items = new Dictionary<byte, IReadOnlyList<string>>();
                var weights = new Dictionary<byte, string>();
                foreach (var entry in slot.GetProperty("values").EnumerateArray())
                {
                    var value = checked((byte)entry.GetProperty("value").GetInt32());
                    if (entry.TryGetProperty("weight", out var weight))
                    {
                        weights[value] = weight.GetString() ?? "";
                    }
                    else if (entry.TryGetProperty("items", out var list))
                    {
                        items[value] = Names(list);
                    }
                }

                slots[slot.GetProperty("name").GetString() ?? ""] = new Slot(offset, tintOffset, items, weights);
            }

            var sets = new Dictionary<string, IReadOnlyList<string>>(StringComparer.Ordinal);
            if (root.TryGetProperty("body_armor", out var bodyArmor) && bodyArmor.TryGetProperty("sets", out var setList))
            {
                foreach (var set in setList.EnumerateArray())
                {
                    var parts = set.GetProperty("parts").EnumerateArray().Select(p => checked((byte)p.GetInt32())).ToArray();
                    if (parts.Length == BodyArmorSlots.Length)
                    {
                        sets[Convert.ToHexString(parts)] = Names(set.GetProperty("items"));
                    }
                }
            }

            var colors = root.TryGetProperty("tints", out var tints) && tints.TryGetProperty("colors", out var colorList)
                ? colorList.EnumerateArray().Select(c => c.GetString() ?? "").ToList()
                : [];

            return new D2EquipmentMap(slots, sets, colors);
        }
        catch (Exception ex) when (ex is JsonException or KeyNotFoundException or InvalidOperationException or OverflowException)
        {
            throw new FormatException("Malformed Diablo II equipment file.", ex);
        }
    }

    /// <summary>
    /// What a character is wearing, in lobby order: helm, body armor, right hand, left hand, shield,
    /// then the Necromancer's shrunken head. Empty for anything that isn't a D2 realm character, and
    /// for a character wearing nothing the map can name (a server that doesn't track items sends
    /// 255 in every slot).
    /// </summary>
    public IReadOnlyList<D2WornItem> Describe(string statString)
    {
        if (D2Character.PortraitBytes(statString) is not { } portrait)
        {
            return [];
        }

        var worn = new List<D2WornItem>();
        AddPart(worn, portrait, "head", "Helm");
        AddBodyArmor(worn, portrait);

        var rightHand = Value(portrait, "right_hand");
        AddPart(worn, portrait, "right_hand", "Right hand");
        // A crossbow repeats in the left hand, and so does an identical second one-hander; either
        // way listing it twice says nothing new.
        if (Value(portrait, "left_hand") != rightHand)
        {
            AddPart(worn, portrait, "left_hand", "Left hand");
        }

        AddPart(worn, portrait, "shield", "Shield");
        AddPart(worn, portrait, "special", "Shrunken head");
        return worn;
    }

    /// <summary>"Wearing: Cap / War Hat / Shako, Dusk Shroud (Crystal Blue), Eldritch Orb" — or "" when there's nothing to say.</summary>
    public string DescribeLine(string statString) =>
        Describe(statString) is { Count: > 0 } worn ? "Wearing: " + string.Join(", ", worn) : "";

    private byte? Value(byte[] portrait, string slotName) =>
        _slots.TryGetValue(slotName, out var slot) ? portrait[slot.Offset] : null;

    private void AddPart(List<D2WornItem> worn, byte[] portrait, string slotName, string label)
    {
        if (!_slots.TryGetValue(slotName, out var slot) || portrait[slot.Offset] == None ||
            !slot.Items.TryGetValue(portrait[slot.Offset], out var looks) || looks.Count == 0)
        {
            return;
        }

        worn.Add(new D2WornItem(label, looks, Tint(portrait[slot.TintOffset])));
    }

    private void AddBodyArmor(List<D2WornItem> worn, byte[] portrait)
    {
        var parts = new byte[BodyArmorSlots.Length];
        for (var i = 0; i < BodyArmorSlots.Length; i++)
        {
            if (!_slots.TryGetValue(BodyArmorSlots[i], out var slot) || portrait[slot.Offset] == None)
            {
                return;
            }

            parts[i] = portrait[slot.Offset];
        }

        var tint = Tint(portrait[_slots["torso"].TintOffset]);
        if (_bodyArmorSets.TryGetValue(Convert.ToHexString(parts), out var armors) && armors.Count > 0)
        {
            worn.Add(new D2WornItem("Armor", armors, tint));
        }
        else if (_slots["torso"].Weights.TryGetValue(parts[0], out var weight))
        {
            worn.Add(new D2WornItem("Armor", [$"{char.ToUpperInvariant(weight[0])}{weight[1..]} armor"], tint));
        }
    }

    /// <summary>A tint byte is transform × 32 + colour + 1; 255 is untinted.</summary>
    private string? Tint(byte tint)
    {
        if (tint == None || tint == 0)
        {
            return null;
        }

        var color = (tint - 1) & 0x1F;
        return color < _colors.Count && _colors[color].Length > 0 ? _colors[color] : null;
    }

    private static List<string> Names(JsonElement items) =>
        items.EnumerateArray()
            .Select(i => i.TryGetProperty("name", out var n) ? n.GetString() ?? "" : "")
            .Where(n => n.Length > 0)
            .Distinct(StringComparer.Ordinal)
            .ToList();
}
