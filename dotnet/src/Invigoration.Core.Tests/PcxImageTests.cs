using Invigoration.Core.Imaging;

namespace Invigoration.Core.Tests;

public class PcxImageTests
{
    // Hand-assembles an 8-bit PCX: 128-byte header, the given pixel stream, then the 0x0C marker
    // and a palette where index i is (i, 255 - i, i / 2) so each index's color is distinguishable.
    private static byte[] BuildPcx(int width, int height, int bytesPerLine, byte[] pixelData, byte encoding = 1)
    {
        var header = new byte[128];
        header[0] = 0x0A;
        header[1] = 5;
        header[2] = encoding;
        header[3] = 8;
        header[8] = (byte)(width - 1);
        header[9] = (byte)((width - 1) >> 8);
        header[10] = (byte)(height - 1);
        header[11] = (byte)((height - 1) >> 8);
        header[65] = 1;
        header[66] = (byte)bytesPerLine;
        header[67] = (byte)(bytesPerLine >> 8);

        var palette = new byte[768];
        for (var i = 0; i < 256; i++)
        {
            palette[i * 3] = (byte)i;
            palette[i * 3 + 1] = (byte)(255 - i);
            palette[i * 3 + 2] = (byte)(i / 2);
        }

        return [.. header, .. pixelData, 0x0C, .. palette];
    }

    private static (byte B, byte G, byte R, byte A) PixelAt(PcxImage image, int x, int y)
    {
        var i = (y * image.Width + x) * 4;
        return (image.Bgra[i], image.Bgra[i + 1], image.Bgra[i + 2], image.Bgra[i + 3]);
    }

    private static (byte B, byte G, byte R, byte A) Color(int index) => ((byte)(index / 2), (byte)(255 - index), (byte)index, 0xFF);

    [Fact]
    public void Decode_ExpandsRunsAndLiteralsThroughThePalette()
    {
        // Row 0: a run of three index-7 pixels then literal index 9. Row 1: literal 1, 2, 3, 4.
        var image = PcxImage.Decode(BuildPcx(4, 2, 4, [0xC3, 7, 9, 1, 2, 3, 4]));

        Assert.Equal(4, image.Width);
        Assert.Equal(2, image.Height);
        Assert.Equal(Color(7), PixelAt(image, 0, 0));
        Assert.Equal(Color(7), PixelAt(image, 2, 0));
        Assert.Equal(Color(9), PixelAt(image, 3, 0));
        Assert.Equal(Color(4), PixelAt(image, 3, 1));
    }

    // A literal byte of 0xC0 or above can't be written bare (it would read as a run), so encoders
    // emit it as a run of one.
    [Fact]
    public void Decode_HandlesRunOfOneForHighIndexes()
    {
        var image = PcxImage.Decode(BuildPcx(2, 1, 2, [0xC1, 0xD0, 0xC1, 0xFF]));

        Assert.Equal(Color(0xD0), PixelAt(image, 0, 0));
        Assert.Equal(Color(0xFF), PixelAt(image, 1, 0));
    }

    [Fact]
    public void Decode_RunsMayCrossScanlines()
    {
        var image = PcxImage.Decode(BuildPcx(3, 2, 3, [0xC6, 42]));

        Assert.Equal(Color(42), PixelAt(image, 0, 0));
        Assert.Equal(Color(42), PixelAt(image, 2, 1));
    }

    // Scanlines are padded to an even byte count; the padding byte isn't a pixel.
    [Fact]
    public void Decode_SkipsScanlinePadding()
    {
        var image = PcxImage.Decode(BuildPcx(3, 2, 4, [10, 11, 12, 0, 20, 21, 22, 0]));

        Assert.Equal(Color(12), PixelAt(image, 2, 0));
        Assert.Equal(Color(20), PixelAt(image, 0, 1));
    }

    [Fact]
    public void Decode_ReadsUncompressedData()
    {
        var image = PcxImage.Decode(BuildPcx(2, 1, 2, [0xC5, 3], encoding: 0));

        // Raw data: 0xC5 is just a pixel, not a run marker.
        Assert.Equal(Color(0xC5), PixelAt(image, 0, 0));
        Assert.Equal(Color(3), PixelAt(image, 1, 0));
    }

    [Fact]
    public void CropBgra_CopiesTheRequestedTile()
    {
        var image = PcxImage.Decode(BuildPcx(4, 2, 4, [1, 2, 3, 4, 5, 6, 7, 8]));

        var tile = image.CropBgra(1, 1, 2, 1);

        Assert.Equal(8, tile.Length);
        Assert.Equal(Color(6), (tile[0], tile[1], tile[2], tile[3]));
        Assert.Equal(Color(7), (tile[4], tile[5], tile[6], tile[7]));
    }

    [Fact]
    public void CropBgra_RejectsRectanglesOutsideTheImage()
    {
        var image = PcxImage.Decode(BuildPcx(4, 2, 4, [1, 2, 3, 4, 5, 6, 7, 8]));

        Assert.Throws<ArgumentOutOfRangeException>(() => image.CropBgra(3, 0, 2, 1));
    }

    [Fact]
    public void Decode_RejectsNonPcxAndUnsupportedVariants()
    {
        Assert.Throws<FormatException>(() => PcxImage.Decode(new byte[1000]));

        var fourBit = BuildPcx(2, 1, 2, [1, 2]);
        fourBit[3] = 4;
        Assert.Throws<FormatException>(() => PcxImage.Decode(fourBit));

        var noPalette = BuildPcx(2, 1, 2, [1, 2]);
        noPalette[^769] = 0;
        Assert.Throws<FormatException>(() => PcxImage.Decode(noPalette));
    }
}
