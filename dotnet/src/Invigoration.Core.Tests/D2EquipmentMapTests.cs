using System.Text;
using Invigoration.Core.StatString;
using static Invigoration.Core.Tests.D2EquipmentTestData;

namespace Invigoration.Core.Tests;

public class D2EquipmentMapTests
{
    private static readonly D2EquipmentMap Map = D2EquipmentMap.Parse(Bytes);

    [Fact]
    public void Describe_NamesEachWornLook_WithItsTint()
    {
        // Crystal Blue is colour 6: tint byte = colour + 1.
        var tints = Enumerable.Repeat((byte)0xFF, 11).ToArray();
        tints[0] = 7;
        var worn = Map.Describe(Statstring(Gear((0, 57), (1, 1), (2, 1), (3, 1), (4, 1), (8, 2), (9, 2), (5, 51), (7, 81)), tints));

        Assert.Collection(worn,
            h => { Assert.Equal("Helm", h.Slot); Assert.Equal(["Cap", "War Hat", "Shako"], h.Looks); Assert.Equal("Crystal Blue", h.Tint); },
            a => { Assert.Equal("Armor", a.Slot); Assert.Equal(["Quilted Armor", "Dusk Shroud"], a.Looks); Assert.Null(a.Tint); },
            r => { Assert.Equal("Right hand", r.Slot); Assert.Equal(4, r.Looks.Count); },
            s => { Assert.Equal("Shield", s.Slot); Assert.Equal(["Kite Shield", "Monarch"], s.Looks); });
    }

    [Fact]
    public void DescribeLine_ReadsNaturally_AndCapsLongLookLists()
    {
        var line = Map.DescribeLine(Statstring(Gear((0, 57), (5, 51))));

        Assert.Equal("Wearing: Cap / War Hat / Shako, Eagle Orb / Glowing Orb / Eldritch Orb / …", line);
    }

    // A body armour pattern the map has no set for still says how heavy it is.
    [Fact]
    public void Describe_UnknownArmorPattern_FallsBackToItsWeight()
    {
        var worn = Map.Describe(Statstring(Gear((1, 3), (2, 3), (3, 3), (4, 3), (8, 3), (9, 3))));

        var armor = Assert.Single(worn);
        Assert.Equal(["Heavy armor"], armor.Looks);
    }

    // A crossbow repeats in the left hand; saying it twice adds nothing.
    [Fact]
    public void Describe_ACrossbowIsListedOnce()
    {
        var worn = Map.Describe(Statstring(Gear((5, 70), (6, 70))));

        Assert.Equal("Right hand", Assert.Single(worn).Slot);
    }

    [Fact]
    public void Describe_ABowInTheLeftHandIsListed()
    {
        Assert.Equal("Left hand", Assert.Single(Map.Describe(Statstring(Gear((6, 41))))).Slot);
    }

    // Command Center's own realm sends 255 in every slot until its game server keeps items.
    [Fact]
    public void Describe_NothingWorn_SaysNothing()
    {
        Assert.Empty(Map.Describe(Statstring(Gear())));
        Assert.Equal("", Map.DescribeLine(Statstring(Gear())));
    }

    [Theory]
    [InlineData("PX2D")]
    [InlineData("VD2DUSEast,Kilua,short")]
    [InlineData("RATS 0 0 7 0 0 0 0 0 RATS")]
    [InlineData("")]
    public void Describe_NotARealmCharacter_IsEmpty(string statString)
    {
        Assert.Empty(Map.Describe(statString));
    }

    [Fact]
    public void Describe_ValuesTheMapDoesntKnow_AreSkipped()
    {
        Assert.Empty(Map.Describe(Statstring(Gear((0, 200), (7, 201)))));
    }

    [Theory]
    [InlineData("\"format\": \"bnetcc-d2-equipment\"", "\"format\": \"something-else\"")]
    [InlineData("\"version\": 1", "\"version\": 2")]
    [InlineData("\"offset\": 2,", "\"offset\": 40,")]
    [InlineData("\"value\": 57,", "\"value\": 300,")]
    public void Parse_RejectsWhatItCantTrust(string find, string replace)
    {
        var json = Json.Replace(find, replace, StringComparison.Ordinal);
        Assert.NotEqual(Json, json);

        Assert.Throws<FormatException>(() => D2EquipmentMap.Parse(Encoding.UTF8.GetBytes(json)));
    }

    [Fact]
    public void Parse_RejectsNonJson()
    {
        Assert.Throws<FormatException>(() => D2EquipmentMap.Parse("not json at all"u8));
        Assert.Throws<FormatException>(() => D2EquipmentMap.Parse("{}"u8));
    }
}
