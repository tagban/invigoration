using Invigoration.Core.StatString;

namespace Invigoration.Core.Chat;

/// <summary>Which of Battle.net's two Diablo II chat-portrait sheets a tile comes from.</summary>
public enum D2PortraitSheet
{
    /// <summary>D2DV.pcx — classic Diablo II: 5 classes, 4 acts per difficulty, dark red tile border.</summary>
    Classic,

    /// <summary>D2XP.pcx — Lord of Destruction: 7 classes, 5 acts per difficulty, gold tile border.</summary>
    Expansion,
}

/// <summary>
/// Picks a Diablo II character's real Battle.net channel-list portrait out of D2DV.pcx/D2XP.pcx
/// (files.bnetdocs.org/Battle.net/Icons/). Both sheets are 28×14 tiles packed with no gutters, laid
/// out as a clean cross-product — decoded 2026-09-12 and checked tile by tile against the art:
/// </summary>
/// <remarks>
/// <para><b>Columns</b> = six status groups × the class count, each group holding every class in
/// order (Amazon, Sorceress, Necromancer, Paladin, Barbarian, then Druid, Assassin): plain,
/// hardcore (red "H"), hardcore dead (skull), then those same three again for ladder (red "L").
/// A softcore character's dead flag has no tile of its own — only hardcore death is permanent.</para>
/// <para><b>Rows</b> = one per act-progression value, plus row 0: rows 1.. carry the act as a green
/// (Normal), yellow (Nightmare) or red (Hell) roman numeral, and the last row is the all-acts-
/// completed sprite tinted gold. So <c>row = progress + 1</c> — D2DV's 14 rows cover progress 0-12,
/// D2XP's 17 cover 0-15. Row 0 (no numeral) is used here for a progress value outside the sheet.</para>
/// </remarks>
public static class D2PortraitTile
{
    public const int TileWidth = 28;
    public const int TileHeight = 14;

    public static int ColumnCount(D2PortraitSheet sheet) => 6 * ClassCount(sheet);

    public static int RowCount(D2PortraitSheet sheet) => sheet == D2PortraitSheet.Expansion ? 17 : 14;

    private static int ClassCount(D2PortraitSheet sheet) =>
        sheet == D2PortraitSheet.Expansion ? D2Character.ExpansionClassCount : D2Character.ClassicClassCount;

    /// <summary>Resolves a statstring straight to its tile. False for non-D2 products, Open characters, and anything whose class the sheet has no column for.</summary>
    public static bool TryGet(string statString, out D2PortraitSheet sheet, out int column, out int row)
    {
        sheet = default;
        column = row = 0;
        return D2Character.TryParse(statString, out var character) && TryGet(character, out sheet, out column, out row);
    }

    public static bool TryGet(D2Character character, out D2PortraitSheet sheet, out int column, out int row)
    {
        sheet = character.Expansion ? D2PortraitSheet.Expansion : D2PortraitSheet.Classic;
        column = row = 0;

        var classCount = ClassCount(sheet);
        if (character.ClassIndex < 0 || character.ClassIndex >= classCount)
        {
            return false;
        }

        var statusGroup = !character.Hardcore ? 0 : character.Dead ? 2 : 1;
        if (character.Ladder)
        {
            statusGroup += 3;
        }

        column = statusGroup * classCount + character.ClassIndex;

        var progressRow = character.Progress + 1;
        row = progressRow < RowCount(sheet) ? progressRow : 0;
        return true;
    }
}
