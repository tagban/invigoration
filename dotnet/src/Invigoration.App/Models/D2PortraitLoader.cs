using Avalonia;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Invigoration.Core.Chat;
using Invigoration.Core.Imaging;

namespace Invigoration.App.Models;

/// <summary>
/// Cuts a Diablo II character's real Battle.net portrait out of the bundled D2DV.pcx/D2XP.pcx
/// sheets (see <see cref="D2PortraitTile"/> for how a statstring picks the tile). Each sheet is
/// decoded once on first use and each tile becomes its own 28×14 bitmap the first time anyone in
/// a channel needs it, then is shared — a full channel of the same class and progress costs one
/// bitmap, not one per row, which matters during a mass-join flood.
/// </summary>
public static class D2PortraitLoader
{
    private static readonly Dictionary<D2PortraitSheet, PcxImage?> Sheets = [];
    private static readonly Dictionary<(D2PortraitSheet Sheet, int Column, int Row), Bitmap?> Tiles = [];

    public static Bitmap? Get(string statString)
    {
        if (!D2PortraitTile.TryGet(statString, out var sheet, out var column, out var row))
        {
            return null;
        }

        var key = (sheet, column, row);
        if (Tiles.TryGetValue(key, out var cached))
        {
            return cached;
        }

        var tile = LoadSheet(sheet) is { } image ? CutTile(image, column, row) : null;
        Tiles[key] = tile;
        return tile;
    }

    private static PcxImage? LoadSheet(D2PortraitSheet sheet)
    {
        if (Sheets.TryGetValue(sheet, out var cached))
        {
            return cached;
        }

        PcxImage? image;
        try
        {
            var name = sheet == D2PortraitSheet.Expansion ? "D2XP" : "D2DV";
            using var stream = AssetLoader.Open(new Uri($"avares://Invigoration.App/Assets/D2Portraits/{name}.pcx"));
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            image = PcxImage.Decode(buffer.ToArray());
        }
        catch (Exception ex) when (ex is FileNotFoundException or FormatException)
        {
            // A missing or damaged sheet just means no portraits; callers fall back to the product icon.
            image = null;
        }

        // A sheet that isn't the expected grid would put every tile in the wrong place — treat it as missing.
        if (image is not null &&
            (image.Width != D2PortraitTile.ColumnCount(sheet) * D2PortraitTile.TileWidth ||
             image.Height != D2PortraitTile.RowCount(sheet) * D2PortraitTile.TileHeight))
        {
            image = null;
        }

        Sheets[sheet] = image;
        return image;
    }

    private static Bitmap CutTile(PcxImage sheet, int column, int row)
    {
        var pixels = sheet.CropBgra(column * D2PortraitTile.TileWidth, row * D2PortraitTile.TileHeight,
            D2PortraitTile.TileWidth, D2PortraitTile.TileHeight);

        var bitmap = new WriteableBitmap(new PixelSize(D2PortraitTile.TileWidth, D2PortraitTile.TileHeight),
            new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Opaque);
        using var framebuffer = bitmap.Lock();
        for (var y = 0; y < D2PortraitTile.TileHeight; y++)
        {
            System.Runtime.InteropServices.Marshal.Copy(pixels, y * D2PortraitTile.TileWidth * 4,
                framebuffer.Address + y * framebuffer.RowBytes, D2PortraitTile.TileWidth * 4);
        }

        return bitmap;
    }
}
