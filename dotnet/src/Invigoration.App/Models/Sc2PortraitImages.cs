using Avalonia;
using Avalonia.Media.Imaging;
using Invigoration.Core.Sc2;

namespace Invigoration.App.Models;

/// <summary>StarCraft II portraits cut from the downloaded sheets (Sc2PortraitStore), once each. Null until the sheets are downloaded.</summary>
public static class Sc2PortraitImages
{
    private static readonly Dictionary<int, Bitmap?> Sheets = [];
    private static readonly Dictionary<(ushort, ushort), Bitmap?> Cells = [];

    static Sc2PortraitImages() => Sc2PortraitStore.Downloaded += () =>
    {
        lock (Cells)
        {
            Sheets.Clear();
            Cells.Clear();
        }
    };

    public static Bitmap? Get(ushort sheet, ushort cell)
    {
        lock (Cells)
        {
            if (Cells.TryGetValue((sheet, cell), out var cached))
            {
                return cached;
            }

            var image = Sheet(sheet) is { } source ? Cut(source, cell) : null;
            if (image is not null)
            {
                Cells[(sheet, cell)] = image;
            }

            return image;
        }
    }

    private static Bitmap? Sheet(int sheet)
    {
        if (Sheets.TryGetValue(sheet, out var loaded))
        {
            return loaded;
        }

        var path = Sc2PortraitStore.SheetPath(sheet);
        if (!File.Exists(path))
        {
            return null;
        }

        try
        {
            using var stream = File.OpenRead(path);
            return Sheets[sheet] = new Bitmap(stream);
        }
        catch (Exception ex) when (ex is IOException or NotSupportedException)
        {
            return Sheets[sheet] = null;
        }
    }

    private static Bitmap Cut(Bitmap sheet, ushort cell)
    {
        var size = Sc2PortraitStore.CellSize;
        var x = cell % Sc2PortraitStore.SheetColumns * size;
        var y = cell / Sc2PortraitStore.SheetColumns * size;
        var image = new RenderTargetBitmap(new PixelSize(size, size));
        using (var context = image.CreateDrawingContext())
        {
            context.DrawImage(sheet, new Rect(x, y, size, size), new Rect(0, 0, size, size));
        }

        return image;
    }
}
