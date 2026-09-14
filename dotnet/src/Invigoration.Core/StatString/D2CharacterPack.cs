using System.IO.Compression;
using System.Text.Json;
using Invigoration.Core.Imaging;

namespace Invigoration.Core.StatString;

/// <summary>One dressed character, ready to play: frames side by side in <see cref="Strip"/> (RGBA, top to bottom), cropped to the figure, feet on the bottom row.</summary>
public sealed record D2CharacterFrames(string Animation, int FrameWidth, int FrameHeight, byte[] Strip, IReadOnlyList<int> FrameDurationsMs);

/// <summary>
/// Command Center's <c>d2-characters.zip</c> (format <c>bnetcc-d2-characters</c>, version 1): the
/// Diablo II character-select animations as layered GIFs, and the rules for stacking them from a
/// realm character's statstring. <see cref="Compose(string)"/> follows the manifest's own rules
/// step by step — docs/D2-CHARACTER-DATA-FOR-BOTS.md explains them. The zip comes from the network,
/// so the manifest is checked as it's read, only entries it names are ever opened, and a part that
/// doesn't match its own description is simply left undrawn.
/// </summary>
public sealed class D2CharacterPack : IDisposable
{
    public const string Format = "bnetcc-d2-characters";
    public const int Version = 1;

    private const int Components = 16;
    private const int MaxEntryBytes = 4 * 1024 * 1024;
    private const int MaxPartFrames = 64;
    private const int MaxCachedParts = 256;

    private readonly ZipArchive _zip;
    private readonly Lock _sync = new();
    private readonly Dictionary<string, GifImage?> _partCache = new(StringComparer.Ordinal);

    private readonly Slot?[] _slots;
    private readonly int[][] _handPairs;
    private readonly string[] _classes;
    private readonly string[] _components;
    private readonly string[] _weaponClasses;
    private readonly Dictionary<string, Animation> _animations;
    private readonly Dictionary<string, Part> _parts;
    private readonly byte[] _palette;
    private readonly byte[] _tints;
    private readonly int _tickMs;

    private D2CharacterPack(ZipArchive zip)
    {
        _zip = zip;
        using var manifest = JsonDocument.Parse(ReadEntry("manifest.json"));
        var root = manifest.RootElement;
        if (!root.TryGetProperty("format", out var format) || format.ValueKind != JsonValueKind.String || format.GetString() != Format ||
            !root.TryGetProperty("version", out var version) || version.ValueKind != JsonValueKind.Number || version.GetInt32() != Version)
        {
            throw new FormatException($"Not a version {Version} {Format} pack.");
        }

        _classes = Strings(root, "classes");
        _components = Strings(root, "components");
        _weaponClasses = Strings(root, "weapon_classes");
        if (_classes.Length < 7 || _components.Length < Components)
        {
            throw new FormatException("The pack's class or component table is short.");
        }

        _handPairs = [.. root.GetProperty("hand_pairs").EnumerateArray().Select(row => row.EnumerateArray().Select(v => v.GetInt32()).ToArray())];

        _slots = new Slot?[256];
        foreach (var slot in root.GetProperty("slots").EnumerateArray())
        {
            var value = slot.GetProperty("value").GetInt32();
            if (value is >= 0 and < 256)
            {
                _slots[value] = new Slot(
                    slot.GetProperty("armor").GetBoolean(),
                    slot.GetProperty("helm").GetBoolean(),
                    slot.TryGetProperty("code", out var code) && code.ValueKind == JsonValueKind.String ? code.GetString() : null,
                    slot.GetProperty("hand").GetInt32(),
                    slot.GetProperty("two_handed").GetInt32(),
                    slot.GetProperty("reserved_hand").GetInt32());
            }
        }

        _animations = new Dictionary<string, Animation>(StringComparer.Ordinal);
        foreach (var entry in root.GetProperty("animations").EnumerateObject())
        {
            var a = entry.Value;
            var layers = a.GetProperty("layers").EnumerateArray()
                .Select(l => (Component: l[0].GetInt32(), WeaponClass: l[1].GetString() ?? ""))
                .Where(l => l.Component is >= 0 and < Components)
                .ToArray();
            var order = a.GetProperty("order").EnumerateArray()
                .Select(f => f.EnumerateArray().Select(c => c.GetInt32()).Where(c => c is >= 0 and < Components).ToArray())
                .ToArray();
            var sequence = a.GetProperty("sequence").EnumerateArray()
                .Select(s => (Frame: s[0].GetInt32(), Ticks: s[1].GetInt32()))
                .Where(s => s.Frame >= 0 && s.Frame < order.Length && s.Ticks is > 0 and <= 1000)
                .ToArray();
            if (sequence.Length > 0)
            {
                _animations[entry.Name.ToUpperInvariant()] = new Animation(layers, order, sequence);
            }
        }

        _parts = new Dictionary<string, Part>(StringComparer.Ordinal);
        foreach (var entry in root.GetProperty("parts").EnumerateObject())
        {
            var p = entry.Value;
            var part = new Part(
                p.GetProperty("file").GetString() ?? "",
                p.GetProperty("frames").GetInt32(),
                p.GetProperty("width").GetInt32(),
                p.GetProperty("height").GetInt32(),
                p.GetProperty("left").GetInt32(),
                p.GetProperty("top").GetInt32());
            if (part.Frames is > 0 and <= MaxPartFrames && part.Width is > 0 and <= 512 && part.Height is > 0 and <= 512 &&
                Math.Abs(part.Left) <= 512 && Math.Abs(part.Top) <= 512)
            {
                _parts[entry.Name.ToUpperInvariant()] = part;
            }
        }

        _tickMs = root.TryGetProperty("tick_ms", out var tick) && tick.TryGetInt32(out var t) && t is > 0 and <= 1000 ? t : 40;
        _palette = ReadEntry(root.GetProperty("palette").GetProperty("file").GetString() ?? "");
        _tints = ReadEntry(root.GetProperty("tints").GetProperty("file").GetString() ?? "");
        if (_palette.Length != 768 || _tints.Length != 8 * 21 * 256)
        {
            throw new FormatException("The pack's palette or tint tables are the wrong size.");
        }
    }

    /// <summary>Opens a stored pack. The file stays open until the pack is disposed.</summary>
    /// <exception cref="FormatException">Not a pack this version can read.</exception>
    public static D2CharacterPack Open(string path) => Open(File.OpenRead(path));

    /// <summary>Opens a pack from a stream it then owns.</summary>
    /// <exception cref="FormatException">Not a pack this version can read.</exception>
    public static D2CharacterPack Open(Stream stream)
    {
        ZipArchive? zip = null;
        try
        {
            zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            return new D2CharacterPack(zip);
        }
        catch (Exception ex) when (ex is InvalidDataException or JsonException or KeyNotFoundException or InvalidOperationException or IndexOutOfRangeException)
        {
            zip?.Dispose();
            stream.Dispose();
            throw new FormatException("The character pack isn't readable.", ex);
        }
        catch
        {
            zip?.Dispose();
            stream.Dispose();
            throw;
        }
    }

    /// <summary>The animation a statstring's character stands in (e.g. "BATNHTH"), or null when there's nothing to draw.</summary>
    public string? AnimationFor(string statString) =>
        D2Character.PortraitBytes(statString) is { } portrait ? PlanFor(portrait)?.Name : null;

    /// <summary>Draws a realm character from their chat statstring. Null for anything the pack can't draw: another game, an Open character, a dead hardcore character, or gear it has no figure for.</summary>
    public D2CharacterFrames? Compose(string statString) =>
        D2Character.PortraitBytes(statString) is { } portrait ? Compose(portrait) : null;

    /// <summary>Draws a character from their 33-byte portrait.</summary>
    public D2CharacterFrames? Compose(byte[] portrait)
    {
        if (PlanFor(portrait) is not { } plan)
        {
            return null;
        }

        lock (_sync)
        {
            var layers = new (Part Part, GifImage Image, byte[]? Tint)?[Components];
            foreach (var (component, part, tint) in plan.Layers)
            {
                if (LoadPart(part) is { } image)
                {
                    layers[component] = (part, image, TintMap(tint));
                }
            }

            // Canvas around the base point big enough for every part, cropped to the figure afterwards.
            var present = layers.Where(l => l is not null).Select(l => l!.Value.Part).ToArray();
            if (present.Length == 0)
            {
                return null;
            }

            var minX = present.Min(p => p.Left);
            var minY = present.Min(p => p.Top);
            var width = present.Max(p => p.Left + p.Width) - minX;
            var height = present.Max(p => p.Top + p.Height) - minY;

            var sequence = plan.Animation.Sequence;
            var canvases = new byte[sequence.Length][];
            for (var step = 0; step < sequence.Length; step++)
            {
                var frame = sequence[step].Frame;
                var canvas = canvases[step] = new byte[width * height * 4];
                foreach (var component in plan.Animation.Order[frame])
                {
                    if (layers[component] is not { } layer)
                    {
                        continue;
                    }

                    var (part, image, tint) = layer;

                    var indices = image.Frames[Math.Min(frame, image.Frames.Count - 1)];
                    var ox = part.Left - minX;
                    var oy = part.Top - minY;
                    for (var y = 0; y < part.Height; y++)
                    {
                        var row = (oy + y) * width;
                        for (var x = 0; x < part.Width; x++)
                        {
                            var index = indices[y * part.Width + x];
                            if (index == 0)
                            {
                                continue;
                            }

                            var colour = (tint is null ? index : tint[index]) * 3;
                            var o = (row + ox + x) * 4;
                            canvas[o] = _palette[colour];
                            canvas[o + 1] = _palette[colour + 1];
                            canvas[o + 2] = _palette[colour + 2];
                            canvas[o + 3] = 255;
                        }
                    }
                }
            }

            return Crop(plan.Name, canvases, width, height, [.. sequence.Select(s => s.Ticks * _tickMs)]);
        }
    }

    public void Dispose() => _zip.Dispose();

    private Plan? PlanFor(ReadOnlySpan<byte> portrait)
    {
        if (portrait.Length < D2Character.PortraitLength)
        {
            return null;
        }

        var characterClass = portrait[13] - 1;
        if (characterClass is < 0 or > 6)
        {
            return null;
        }

        Span<int> gear = stackalloc int[Components];
        Span<int> tints = stackalloc int[Components];
        gear.Fill(255);
        tints.Fill(255);
        for (var c = 0; c <= 10; c++)
        {
            gear[c] = portrait[2 + c];
            tints[c] = portrait[14 + c];
        }

        // Stance: a dead hardcore character isn't in the pack; hardcore stands in NU, everyone else TN.
        var status = portrait[26];
        if ((status & 0x0C) == 0x0C)
        {
            return null;
        }

        var mode = (status & 0x04) != 0 ? "NU" : "TN";

        // Hands → weapon class.
        int rh = gear[5], lh = gear[6], sh = gear[7];
        var both = rh != 255 && lh != 255;
        var assassin = characterClass == 6;
        int HandClass(int v, bool right)
        {
            if (v == 255 || _slots[v] is not { } s)
            {
                return 0;
            }

            var k = both || (right && lh == 255 && sh == 255 && s.TwoHanded != s.Hand) ? s.TwoHanded : s.Hand;
            if ((k is 13 or 14 && !assassin) || s.Armor)
            {
                k = s.ReservedHand;
            }

            return k is 13 or 14 && !assassin ? 0 : k;
        }

        var right = HandClass(rh, right: true);
        var left = HandClass(lh, right: false);
        if (right < 0 || right >= _handPairs.Length || left < 0 || left >= _handPairs[right].Length)
        {
            return null;
        }

        var weaponClass = _handPairs[right][left];
        if (weaponClass <= 0 || weaponClass >= _weaponClasses.Length)
        {
            return null;
        }

        var classToken = _classes[characterClass];
        var name = (classToken + mode + _weaponClasses[weaponClass]).ToUpperInvariant();
        if (!_animations.TryGetValue(name, out var animation))
        {
            return null;
        }

        var layers = new List<(int Component, Part Part, int Tint)>();
        foreach (var (component, layerWeaponClass) in animation.Layers)
        {
            var v = gear[component];
            var slot = v is 0 or 255 ? null : _slots[v];
            var code = slot?.Code is not { } slotCode || (component == 0 && !slot.Helm) ? "lit" : slotCode;
            var key = (classToken + _components[component] + code + mode + layerWeaponClass).ToUpperInvariant();
            if (_parts.TryGetValue(key, out var part))
            {
                layers.Add((component, part, tints[component]));
            }
        }

        return new Plan(name, animation, layers);
    }

    private byte[]? TintMap(int x)
    {
        if (x is 0 or 255)
        {
            return null;
        }

        var s = x - 1;
        var colour = s & 31;
        var transform = s >> 5;
        if (transform == 0)
        {
            transform = 8;
        }

        if (transform is 3 or 4 || colour > 20)
        {
            return null;
        }

        return _tints.AsSpan(((transform - 1) * 21 + colour) * 256, 256).ToArray();
    }

    /// <summary>A part's decoded GIF, or null if it's missing or doesn't match the manifest. Caller holds _sync.</summary>
    private GifImage? LoadPart(Part part)
    {
        if (_partCache.TryGetValue(part.File, out var cached))
        {
            return cached;
        }

        GifImage? image;
        try
        {
            image = GifImage.Decode(ReadEntry(part.File), MaxPartFrames);
            if (image.Width != part.Width || image.Height != part.Height)
            {
                image = null;
            }
        }
        catch (Exception ex) when (ex is FormatException or InvalidDataException or IOException)
        {
            image = null;
        }

        if (_partCache.Count >= MaxCachedParts)
        {
            _partCache.Clear();
        }

        _partCache[part.File] = image;
        return image;
    }

    /// <summary>Crops every frame to the box that holds the figure in any of them, bottom row the lowest pixel, and lays them side by side.</summary>
    private static D2CharacterFrames? Crop(string name, byte[][] canvases, int width, int height, IReadOnlyList<int> durations)
    {
        int left = width, top = height, right = -1, bottom = -1;
        foreach (var canvas in canvases)
        {
            for (var y = 0; y < height; y++)
            {
                for (var x = 0; x < width; x++)
                {
                    if (canvas[(y * width + x) * 4 + 3] != 0)
                    {
                        left = Math.Min(left, x);
                        right = Math.Max(right, x);
                        top = Math.Min(top, y);
                        bottom = Math.Max(bottom, y);
                    }
                }
            }
        }

        if (right < 0)
        {
            return null;
        }

        var frameWidth = right - left + 1;
        var frameHeight = bottom - top + 1;
        var strip = new byte[frameWidth * canvases.Length * frameHeight * 4];
        for (var f = 0; f < canvases.Length; f++)
        {
            for (var y = 0; y < frameHeight; y++)
            {
                Buffer.BlockCopy(canvases[f], ((top + y) * width + left) * 4, strip, (y * frameWidth * canvases.Length + f * frameWidth) * 4, frameWidth * 4);
            }
        }

        return new D2CharacterFrames(name, frameWidth, frameHeight, strip, durations);
    }

    /// <summary>Reads a zip entry the manifest names — never extracted, and bounded in size.</summary>
    private byte[] ReadEntry(string name)
    {
        var entry = _zip.GetEntry(name) ?? throw new FormatException($"The pack has no {name}.");
        if (entry.Length is < 0 or > MaxEntryBytes)
        {
            throw new FormatException($"{name} is too large.");
        }

        using var stream = entry.Open();
        var bytes = new byte[entry.Length];
        stream.ReadExactly(bytes);
        return bytes;
    }

    private static string[] Strings(JsonElement root, string property) =>
        [.. root.GetProperty(property).EnumerateArray().Select(e => e.GetString() ?? "")];

    private sealed record Slot(bool Armor, bool Helm, string? Code, int Hand, int TwoHanded, int ReservedHand);

    private sealed record Part(string File, int Frames, int Width, int Height, int Left, int Top);

    private sealed record Animation((int Component, string WeaponClass)[] Layers, int[][] Order, (int Frame, int Ticks)[] Sequence);

    private sealed record Plan(string Name, Animation Animation, List<(int Component, Part Part, int Tint)> Layers);
}
