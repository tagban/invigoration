using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using Invigoration.Core.StatString;

namespace Invigoration.Core.Tests;

/// <summary>
/// The d2-characters.zip rules (docs/D2-CHARACTER-DATA-FOR-BOTS.md), on a tiny made-up pack: a
/// Barbarian whose torso, head and right hand are a few pixels each, so every rule shows up as a
/// specific palette index at a specific pixel. Palette index i is the colour (i, 2i, 255 - i).
/// </summary>
public class D2CharacterPackTests
{
    private const byte T = 0; // transparent

    private static JsonObject Manifest() => new()
    {
        ["format"] = "bnetcc-d2-characters",
        ["version"] = 1,
        ["tick_ms"] = 40,
        ["classes"] = new JsonArray("AM", "SO", "NE", "PA", "BA", "DZ", "AI"),
        ["components"] = new JsonArray("HD", "TR", "LG", "RA", "LA", "RH", "LH", "SH", "S1", "S2", "S3", "S4", "S5", "S6", "S7", "S8"),
        ["weapon_classes"] = new JsonArray("", "hth", "1hs"),
        // [right hand class][left hand class]: empty hands → hth, a one-hander in the right → 1hs, anything in the left alone → nothing.
        ["hand_pairs"] = new JsonArray(new JsonArray(1, 0, 0), new JsonArray(1, 0, 0), new JsonArray(2, 0, 0)),
        ["slots"] = new JsonArray(
            Slot(0, null),
            Slot(1, "lit", armor: true),
            Slot(2, "med", armor: true),
            Slot(17, "ssd", hand: 2),
            Slot(18, "zzz", hand: 2),
            Slot(41, "sbw", hand: 1),
            Slot(57, "cap", helm: true),
            Slot(60, "xxx")),
        ["animations"] = new JsonObject
        {
            ["BATNHTH"] = Animation([[1, "hth"], [0, "hth"]], [[1, 0], [0, 1]], [[0, 2], [1, 3]]),
            ["BANUHTH"] = Animation([[1, "hth"]], [[1]], [[0, 1]]),
            ["BATN1HS"] = Animation([[1, "hth"], [5, "1hs"]], [[1, 5]], [[0, 1]]),
        },
        ["parts"] = new JsonObject
        {
            ["BATRLITTNHTH"] = Part("BATRLITTNHTH", frames: 2, width: 2, height: 3, left: -1, top: -3),
            ["BATRMEDTNHTH"] = Part("BATRMEDTNHTH", frames: 1, width: 2, height: 3, left: -1, top: -3),
            ["BAHDLITTNHTH"] = Part("BAHDLITTNHTH", frames: 1, width: 2, height: 2, left: 0, top: -3),
            ["BAHDCAPTNHTH"] = Part("BAHDCAPTNHTH", frames: 1, width: 2, height: 2, left: 0, top: -3),
            ["BATRLITNUHTH"] = Part("BATRLITNUHTH", frames: 1, width: 1, height: 1, left: 0, top: -1),
            ["BARHSSDTN1HS"] = Part("BARHSSDTN1HS", frames: 1, width: 1, height: 1, left: 1, top: -1),
        },
        ["palette"] = new JsonObject { ["file"] = "palette.bin" },
        ["tints"] = new JsonObject { ["file"] = "tints.bin", ["colours"] = 21, ["transforms"] = 8 },
    };

    private static readonly Dictionary<string, byte[][]> PartPixels = new()
    {
        // Torso: two frames, bottom row empty so the figure gets cropped.
        ["BATRLITTNHTH"] = [[10, 10, 10, 10, T, T], [11, 11, 11, 11, T, T]],
        ["BATRMEDTNHTH"] = [[20, 20, 20, 20, T, T]],
        ["BAHDLITTNHTH"] = [[30, T, 30, 30]],
        ["BAHDCAPTNHTH"] = [[40, T, 40, 40]],
        ["BATRLITNUHTH"] = [[50]],
        ["BARHSSDTN1HS"] = [[60]],
    };

    private static JsonObject Slot(int value, string? code, bool armor = false, bool helm = false, int hand = 0) => new()
    {
        ["value"] = value,
        ["code"] = code,
        ["armor"] = armor,
        ["helm"] = helm,
        ["hand"] = hand,
        ["two_handed"] = hand,
        ["reserved_hand"] = 0,
    };

    private static JsonObject Animation(object[][] layers, int[][] order, int[][] sequence) => new()
    {
        ["layers"] = new JsonArray([.. layers.Select(l => (JsonNode)new JsonArray((int)l[0], (string)l[1]))]),
        ["order"] = new JsonArray([.. order.Select(o => (JsonNode)new JsonArray([.. o.Select(c => (JsonNode)c)]))]),
        ["sequence"] = new JsonArray([.. sequence.Select(s => (JsonNode)new JsonArray(s[0], s[1]))]),
    };

    private static JsonObject Part(string name, int frames, int width, int height, int left, int top) => new()
    {
        ["file"] = $"parts/BA/{name}.gif",
        ["frames"] = frames,
        ["width"] = width,
        ["height"] = height,
        ["left"] = left,
        ["top"] = top,
    };

    private static byte[] BuildZip(JsonObject manifest, Action<Dictionary<string, byte[]>>? changeEntries = null)
    {
        var palette = new byte[768];
        for (var i = 0; i < 256; i++)
        {
            palette[i * 3] = (byte)i;
            palette[i * 3 + 1] = (byte)(i * 2);
            palette[i * 3 + 2] = (byte)(255 - i);
        }

        // Identity maps, except transform 1 / colour 5 (tint byte 38), which adds 100.
        var tints = new byte[8 * 21 * 256];
        for (var map = 0; map < 8 * 21; map++)
        {
            for (var p = 0; p < 256; p++)
            {
                tints[map * 256 + p] = (byte)(map == 5 ? (p + 100) % 256 : p);
            }
        }

        var entries = new Dictionary<string, byte[]>
        {
            ["manifest.json"] = Encoding.UTF8.GetBytes(manifest.ToJsonString()),
            ["palette.bin"] = palette,
            ["tints.bin"] = tints,
        };
        foreach (var (name, frames) in PartPixels)
        {
            var part = manifest["parts"]![name]!;
            entries[$"parts/BA/{name}.gif"] = TestGif.Build(part["width"]!.GetValue<int>(), part["height"]!.GetValue<int>(), frames);
        }

        changeEntries?.Invoke(entries);

        var buffer = new MemoryStream();
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, bytes) in entries)
            {
                using var stream = zip.CreateEntry(name, CompressionLevel.NoCompression).Open();
                stream.Write(bytes);
            }
        }

        return buffer.ToArray();
    }

    private static D2CharacterPack OpenPack(JsonObject? manifest = null, Action<Dictionary<string, byte[]>>? changeEntries = null) =>
        D2CharacterPack.Open(new MemoryStream(BuildZip(manifest ?? Manifest(), changeEntries)));

    /// <summary>A realm Barbarian's statstring with the given slots (head, torso, ..., right hand at 5) and tints.</summary>
    private static string Barbarian(Dictionary<int, byte>? gear = null, Dictionary<int, byte>? tints = null, byte status = 0x20, byte characterClass = 5)
    {
        var p = Enumerable.Repeat((byte)0xFF, D2Character.PortraitLength).ToArray();
        p[0] = 0x84;
        p[1] = 0x80;
        p[13] = characterClass;
        p[25] = 30;
        p[26] = status;
        p[27] = 0x80;
        foreach (var (c, v) in gear ?? [])
        {
            p[2 + c] = v;
        }

        foreach (var (c, v) in tints ?? [])
        {
            p[14 + c] = v;
        }

        return "PX2DUSEast,Kilua," + Encoding.Latin1.GetString(p);
    }

    /// <summary>The palette index a composed pixel came from (the test palette's red channel), or null if it's clear.</summary>
    private static byte? IndexAt(D2CharacterFrames frames, int frame, int x, int y)
    {
        var o = (y * frames.FrameWidth * frames.FrameDurationsMs.Count + frame * frames.FrameWidth + x) * 4;
        if (frames.Strip[o + 3] == 0)
        {
            return null;
        }

        Assert.Equal((byte)(frames.Strip[o] * 2), frames.Strip[o + 1]);
        Assert.Equal((byte)(255 - frames.Strip[o]), frames.Strip[o + 2]);
        return frames.Strip[o];
    }

    private static byte?[] Row(D2CharacterFrames frames, int frame, int y) =>
        [.. Enumerable.Range(0, frames.FrameWidth).Select(x => IndexAt(frames, frame, x, y))];

    [Fact]
    public void Compose_DrawsTheBaseLookInDrawOrderCroppedToTheFigure()
    {
        using var pack = OpenPack();

        var frames = pack.Compose(Barbarian());

        Assert.NotNull(frames);
        Assert.Equal("BATNHTH", frames.Animation);
        Assert.Equal(3, frames.FrameWidth);
        Assert.Equal(2, frames.FrameHeight); // the torso's empty bottom row is cropped away
        Assert.Equal([80, 120], frames.FrameDurationsMs);

        // Frame 0 draws the torso, then the head over it; frame 1 the other way round, with the
        // torso's second frame and the one-frame head held on its last.
        Assert.Equal([10, 30, null], Row(frames, 0, 0));
        Assert.Equal([10, 30, 30], Row(frames, 0, 1));
        Assert.Equal([11, 11, null], Row(frames, 1, 0));
        Assert.Equal([11, 11, 30], Row(frames, 1, 1));
    }

    [Fact]
    public void Compose_HardcoreStandsInTheNeutralStance_AndDeadHardcoreDrawsNothing()
    {
        using var pack = OpenPack();

        Assert.Equal("BANUHTH", pack.Compose(Barbarian(status: 0x24))?.Animation);
        Assert.Equal("BANUHTH", pack.AnimationFor(Barbarian(status: 0x24)));
        Assert.Null(pack.Compose(Barbarian(status: 0x2C)));

        // Dead but softcore is just a character.
        Assert.Equal("BATNHTH", pack.AnimationFor(Barbarian(status: 0x28)));
    }

    [Fact]
    public void Compose_DrawsGearByItsCode_ButOnlyARealHelmOnTheHead()
    {
        using var pack = OpenPack();

        var dressed = pack.Compose(Barbarian(gear: new() { [0] = 57, [1] = 2 }))!;
        Assert.Equal([20, 40, null], Row(dressed, 0, 0));

        // Value 60 isn't a helm: the head keeps its base look.
        var notAHelm = pack.Compose(Barbarian(gear: new() { [0] = 60 }))!;
        Assert.Equal([10, 30, null], Row(notAHelm, 0, 0));
    }

    [Fact]
    public void Compose_TintsThroughTheTintTables()
    {
        using var pack = OpenPack();

        // Tint byte 38 → transform 1, colour 5: that table adds 100 to every index it maps.
        var tinted = pack.Compose(Barbarian(tints: new() { [1] = 38 }))!;
        Assert.Equal([110, 30, null], Row(tinted, 0, 0));

        // Transforms 3 and 4 aren't drawn tinted: tint byte 97 is transform 3.
        var untinted = pack.Compose(Barbarian(tints: new() { [1] = 97 }))!;
        Assert.Equal([10, 30, null], Row(untinted, 0, 0));
    }

    [Fact]
    public void Compose_PicksTheWeaponClassFromTheHands()
    {
        using var pack = OpenPack();

        var armed = pack.Compose(Barbarian(gear: new() { [5] = 17 }))!;
        Assert.Equal("BATN1HS", armed.Animation);
        Assert.Contains((byte?)60, Enumerable.Range(0, armed.FrameHeight).SelectMany(y => Row(armed, 0, y)));

        // Something only in the left hand makes weapon class 0: the game draws a fallback there.
        Assert.Null(pack.Compose(Barbarian(gear: new() { [6] = 41 })));
    }

    [Fact]
    public void Compose_LeavesOutPartsThePackDoesntHaveOrThatDontMatchTheirDescription()
    {
        // Gear with no part in the pack: the rest of the character is still drawn.
        using (var pack = OpenPack())
        {
            var frames = pack.Compose(Barbarian(gear: new() { [5] = 18 }))!;
            Assert.Equal("BATN1HS", frames.Animation);
            Assert.DoesNotContain((byte?)60, Enumerable.Range(0, frames.FrameHeight).SelectMany(y => Row(frames, 0, y)));
        }

        // A head GIF that isn't the size the manifest says: left out rather than misdrawn.
        using (var pack = OpenPack(changeEntries: e => e["parts/BA/BAHDLITTNHTH.gif"] = TestGif.Build(3, 3, new byte[9])))
        {
            var frames = pack.Compose(Barbarian())!;
            Assert.Equal(2, frames.FrameWidth);
            Assert.Equal([10, 10], Row(frames, 0, 0));
        }
    }

    [Theory]
    [InlineData("PX2D")] // an Open character
    [InlineData("RATS 0 0 0 0 0")]
    [InlineData("")]
    public void Compose_DrawsNothingForAnythingButARealmCharacter(string statString)
    {
        using var pack = OpenPack();

        Assert.Null(pack.Compose(statString));
    }

    [Fact]
    public void Compose_DrawsNothingForAnUnknownClass()
    {
        using var pack = OpenPack();

        Assert.Null(pack.Compose(Barbarian(characterClass: 9)));
    }

    [Fact]
    public void Open_RejectsPacksItCantRead()
    {
        var newer = Manifest();
        newer["version"] = 2;
        Assert.Throws<FormatException>(() => OpenPack(newer));

        var other = Manifest();
        other["format"] = "something-else";
        Assert.Throws<FormatException>(() => OpenPack(other));

        Assert.Throws<FormatException>(() => OpenPack(changeEntries: e => e["palette.bin"] = new byte[10]));
        Assert.Throws<FormatException>(() => OpenPack(changeEntries: e => e.Remove("manifest.json")));
        Assert.Throws<FormatException>(() => OpenPack(changeEntries: e => e["manifest.json"] = "{ not json"u8.ToArray()));
        Assert.Throws<FormatException>(() => D2CharacterPack.Open(new MemoryStream("not a zip"u8.ToArray())));
    }
}
