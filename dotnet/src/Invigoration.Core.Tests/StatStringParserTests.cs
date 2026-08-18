using Invigoration.Core.StatString;

namespace Invigoration.Core.Tests;

public class StatStringParserTests
{
    [Fact]
    public void Parse_EmptyD2Character_ReturnsOpenCharacter()
    {
        var result = StatStringParser.Parse("VD2D");

        Assert.Equal("Diablo II: (Open Character).", result);
    }

    [Fact]
    public void Parse_D2Character_NonHardcoreNonExpansion_ProducesLevelAndClass()
    {
        var p = new char[28];
        p[13] = (char)1; // charclass raw = 1 -> charclass 0 (amazon)
        p[25] = (char)42; // level
        p[26] = (char)0; // not hardcore, not dead
        p[27] = (char)0; // no title tier

        var stats = "VD2D" + "USEast," + "MyChar," + new string(p);

        var result = StatStringParser.Parse(stats);

        Assert.Equal("Diablo II: (MyChar a level 42 amazon on realm USEast).", result);
    }

    [Fact]
    public void Parse_D2Character_HardcoreDead_MentionsDead()
    {
        var p = new char[28];
        p[13] = (char)2; // charclass raw = 2 -> charclass 1 (sorceress, female)
        p[25] = (char)10;
        p[26] = (char)(0x4 | 0x8); // hardcore + dead
        p[27] = (char)0;

        var stats = "VD2D" + "USWest," + "DeadChar," + new string(p);

        var result = StatStringParser.Parse(stats);

        Assert.Equal("Diablo II: (DeadChar a dead hardcore level 10 sorceress on realm USWest).", result);
    }

    [Fact]
    public void Parse_D2ExpansionCharacter_LadderTitle_UsesExpansionTitles()
    {
        var p = new char[28];
        p[13] = (char)4; // charclass raw 4 -> charclass 3 (paladin, male)
        p[25] = (char)99;
        p[26] = (char)0x20; // ladder flag set (expansion title path)
        p[27] = (char)(1 << 3); // tier 1 -> "Slayer" (not hardcore)

        var stats = "PX2D" + "Europe," + "Hero," + new string(p);

        var result = StatStringParser.Parse(stats);

        Assert.Equal("Diablo II Lord of Destruction: (Slayer Hero a level 99 paladin on realm Europe).", result);
    }

    [Fact]
    public void Parse_UnknownProduct_ReturnsEmpty()
    {
        Assert.Equal("", StatStringParser.Parse("ZZZZ"));
    }
}
