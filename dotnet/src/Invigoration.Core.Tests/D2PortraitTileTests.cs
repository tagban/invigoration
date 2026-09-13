using Invigoration.Core.Chat;
using Invigoration.Core.StatString;

namespace Invigoration.Core.Tests;

public class D2PortraitTileTests
{
    // Builds a real-shaped 33-byte struct: every byte non-zero like the wire's C string, 0xFF for
    // "empty", with the fields the portrait depends on filled in.
    private static string Stats(string product, int classRaw, int flags = 0x80, int actByte = 0x80, int ladder = 0xFF, int level = 42)
    {
        var p = Enumerable.Repeat((char)0xFF, 33).ToArray();
        p[0] = (char)0x84;
        p[1] = (char)0x80;
        p[13] = (char)classRaw;
        p[25] = (char)level;
        p[26] = (char)flags;
        p[27] = (char)actByte;
        p[30] = (char)ladder;
        return product + "USEast," + "Hero," + new string(p);
    }

    [Fact]
    public void TryParse_ReadsEveryPortraitField()
    {
        Assert.True(D2Character.TryParse(Stats("PX2D", classRaw: 7, flags: 0x80 | 0x20 | 0x04 | 0x08, actByte: 0x94, ladder: 0x01, level: 88), out var c));

        Assert.Equal("USEast", c.Realm);
        Assert.Equal("Hero", c.Name);
        Assert.Equal(6, c.ClassIndex);
        Assert.Equal(88, c.Level);
        Assert.True(c.Hardcore);
        Assert.True(c.Dead);
        Assert.True(c.Expansion);
        Assert.True(c.Ladder);
        Assert.Equal(10, c.Progress);
    }

    [Theory]
    [InlineData("VD2D")]
    [InlineData("PX2D")]
    [InlineData("PX2DUSEast")]
    [InlineData("PX2DUSEast,Hero")]
    [InlineData("RATS 0 0 5 0 0 0 0 0 RATS")]
    public void TryParse_RejectsOpenCharactersAndOtherProducts(string statString)
    {
        Assert.False(D2Character.TryParse(statString, out _));
        Assert.False(D2PortraitTile.TryGet(statString, out _, out _, out _));
    }

    // The expansion flag only means anything on the expansion product.
    [Fact]
    public void TryParse_IgnoresTheExpansionFlagOnClassicDiabloII()
    {
        Assert.True(D2Character.TryParse(Stats("VD2D", classRaw: 1, flags: 0x80 | 0x20), out var c));
        Assert.False(c.Expansion);
    }

    // Every documented "Current Act" value (bnetdocs chat-statstrings) maps to its row: row 0 is
    // reserved, so the first act of Normal is row 1 and "all acts completed" is the sheet's last row.
    [Theory]
    [InlineData(0x80, 1)]
    [InlineData(0x86, 4)]
    [InlineData(0x88, 5)]
    [InlineData(0x90, 9)]
    [InlineData(0x96, 12)]
    [InlineData(0x98, 13)]
    public void TryGet_ClassicActByteSelectsRow(int actByte, int expectedRow)
    {
        Assert.True(D2PortraitTile.TryGet(Stats("VD2D", classRaw: 1, actByte: actByte), out var sheet, out _, out var row));
        Assert.Equal(D2PortraitSheet.Classic, sheet);
        Assert.Equal(expectedRow, row);
        Assert.True(row < D2PortraitTile.RowCount(sheet));
    }

    [Theory]
    [InlineData(0x80, 1)]
    [InlineData(0x86, 4)]
    [InlineData(0x8A, 6)]
    [InlineData(0x90, 9)]
    [InlineData(0x94, 11)]
    [InlineData(0x9A, 14)]
    [InlineData(0x9E, 16)]
    public void TryGet_ExpansionActByteSelectsRow(int actByte, int expectedRow)
    {
        Assert.True(D2PortraitTile.TryGet(Stats("PX2D", classRaw: 1, flags: 0xA0, actByte: actByte), out var sheet, out _, out var row));
        Assert.Equal(D2PortraitSheet.Expansion, sheet);
        Assert.Equal(expectedRow, row);
        Assert.True(row < D2PortraitTile.RowCount(sheet));
    }

    [Fact]
    public void TryGet_ProgressPastTheSheetFallsBackToRowZero()
    {
        // 0xFF would be progress 31 — not a real value, and far past either sheet.
        Assert.True(D2PortraitTile.TryGet(Stats("VD2D", classRaw: 1, actByte: 0xFF), out _, out _, out var row));
        Assert.Equal(0, row);
    }

    [Theory]
    // Classic: 5 classes per status group.
    [InlineData("VD2D", 1, 0x80, 0xFF, 0)]    // plain Amazon
    [InlineData("VD2D", 5, 0x80, 0xFF, 4)]    // plain Barbarian
    [InlineData("VD2D", 2, 0x84, 0xFF, 6)]    // hardcore Sorceress
    [InlineData("VD2D", 3, 0x8C, 0xFF, 12)]   // hardcore dead Necromancer
    [InlineData("VD2D", 4, 0x80, 0x01, 18)]   // ladder Paladin
    [InlineData("VD2D", 1, 0x84, 0x01, 20)]   // hardcore ladder Amazon
    [InlineData("VD2D", 5, 0x8C, 0x01, 29)]   // hardcore dead ladder Barbarian — last column
    // Softcore death isn't permanent, so it has no skull tile.
    [InlineData("VD2D", 1, 0x88, 0xFF, 0)]
    // Expansion: 7 classes per status group.
    [InlineData("PX2D", 6, 0xA0, 0xFF, 5)]    // plain Druid
    [InlineData("PX2D", 7, 0xA4, 0xFF, 13)]   // hardcore Assassin
    [InlineData("PX2D", 1, 0xAC, 0xFF, 14)]   // hardcore dead Amazon
    [InlineData("PX2D", 1, 0xA0, 0x01, 21)]   // ladder Amazon
    [InlineData("PX2D", 7, 0xAC, 0x01, 41)]   // hardcore dead ladder Assassin — last column
    public void TryGet_ColumnIsStatusGroupTimesClassCountPlusClass(string product, int classRaw, int flags, int ladder, int expectedColumn)
    {
        Assert.True(D2PortraitTile.TryGet(Stats(product, classRaw, flags, ladder: ladder), out var sheet, out var column, out _));
        Assert.Equal(expectedColumn, column);
        Assert.True(column < D2PortraitTile.ColumnCount(sheet));
    }

    [Theory]
    [InlineData("VD2D", 6, 0x80)]  // a Druid on the classic sheet has no column
    [InlineData("VD2D", 0, 0x80)]  // class byte 0 isn't a class
    [InlineData("PX2D", 8, 0xA0)]  // past Assassin
    public void TryGet_RejectsClassesTheSheetHasNoColumnFor(string product, int classRaw, int flags)
    {
        Assert.False(D2PortraitTile.TryGet(Stats(product, classRaw, flags), out _, out _, out _));
    }

    // Both sheets are exactly this many tiles — matches the real D2DV.pcx (840×196) and D2XP.pcx (1176×238).
    [Theory]
    [InlineData(D2PortraitSheet.Classic, 840, 196)]
    [InlineData(D2PortraitSheet.Expansion, 1176, 238)]
    public void SheetDimensionsMatchTheRealFiles(D2PortraitSheet sheet, int width, int height)
    {
        Assert.Equal(width, D2PortraitTile.ColumnCount(sheet) * D2PortraitTile.TileWidth);
        Assert.Equal(height, D2PortraitTile.RowCount(sheet) * D2PortraitTile.TileHeight);
    }
}
