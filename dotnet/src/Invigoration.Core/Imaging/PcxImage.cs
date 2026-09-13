namespace Invigoration.Core.Imaging;

/// <summary>
/// Minimal decoder for 8-bit palettized PCX — the format of every classic Battle.net icon sheet
/// (files.bnetdocs.org/Battle.net/Icons/*.pcx) and of the ad banners SID_CHECKAD points at. Only
/// what those files use: one 8-bit plane, RLE or raw, with the 256-color VGA palette appended to
/// the end of the file. Pure byte work, no UI types, so the app wraps the pixels however its
/// imaging stack wants.
/// </summary>
public sealed class PcxImage
{
    private const int HeaderSize = 128;
    private const int PaletteSize = 768;

    public int Width { get; }

    public int Height { get; }

    /// <summary>Fully opaque pixels, 4 bytes each in B, G, R, A order, row by row from the top.</summary>
    public byte[] Bgra { get; }

    private PcxImage(int width, int height, byte[] bgra)
    {
        Width = width;
        Height = height;
        Bgra = bgra;
    }

    /// <exception cref="FormatException">Not a PCX, or a PCX variant these sheets don't use.</exception>
    public static PcxImage Decode(ReadOnlySpan<byte> data)
    {
        if (data.Length < HeaderSize + 1 + PaletteSize || data[0] != 0x0A)
        {
            throw new FormatException("Not a PCX file.");
        }

        var encoding = data[2];
        var bitsPerPixel = data[3];
        var planes = data[65];
        if (bitsPerPixel != 8 || planes != 1 || encoding > 1)
        {
            throw new FormatException($"Unsupported PCX variant ({bitsPerPixel} bpp, {planes} planes, encoding {encoding}).");
        }

        var width = ReadUInt16(data, 8) - ReadUInt16(data, 4) + 1;
        var height = ReadUInt16(data, 10) - ReadUInt16(data, 6) + 1;
        var bytesPerLine = ReadUInt16(data, 66);
        if (width <= 0 || height <= 0 || bytesPerLine < width)
        {
            throw new FormatException("PCX header has an invalid size.");
        }

        var paletteStart = data.Length - PaletteSize;
        if (data[paletteStart - 1] != 0x0C)
        {
            throw new FormatException("PCX has no 256-color palette.");
        }

        var palette = data[paletteStart..];
        var scanlines = DecodeScanlines(data[HeaderSize..(paletteStart - 1)], encoding == 1, bytesPerLine * height);

        var bgra = new byte[width * height * 4];
        for (var y = 0; y < height; y++)
        {
            var line = scanlines.AsSpan(y * bytesPerLine, width);
            var outRow = bgra.AsSpan(y * width * 4, width * 4);
            for (var x = 0; x < width; x++)
            {
                var p = line[x] * 3;
                outRow[x * 4] = palette[p + 2];
                outRow[x * 4 + 1] = palette[p + 1];
                outRow[x * 4 + 2] = palette[p];
                outRow[x * 4 + 3] = 0xFF;
            }
        }

        return new PcxImage(width, height, bgra);
    }

    /// <summary>Copies one rectangle out as its own BGRA buffer — how the icon sheets get cut into tiles.</summary>
    public byte[] CropBgra(int x, int y, int width, int height)
    {
        if (x < 0 || y < 0 || width <= 0 || height <= 0 || x + width > Width || y + height > Height)
        {
            throw new ArgumentOutOfRangeException(nameof(x), "Crop rectangle falls outside the image.");
        }

        var result = new byte[width * height * 4];
        for (var row = 0; row < height; row++)
        {
            Bgra.AsSpan(((y + row) * Width + x) * 4, width * 4).CopyTo(result.AsSpan(row * width * 4));
        }

        return result;
    }

    // RLE runs are allowed to cross scanline boundaries, so the whole image decodes as one stream.
    // A truncated stream leaves the remainder as palette index 0 rather than failing the image.
    private static byte[] DecodeScanlines(ReadOnlySpan<byte> encoded, bool rle, int length)
    {
        var output = new byte[length];
        if (!rle)
        {
            encoded[..Math.Min(encoded.Length, length)].CopyTo(output);
            return output;
        }

        var o = 0;
        var i = 0;
        while (o < length && i < encoded.Length)
        {
            var b = encoded[i++];
            if ((b & 0xC0) == 0xC0)
            {
                if (i >= encoded.Length)
                {
                    break;
                }

                var count = Math.Min(b & 0x3F, length - o);
                output.AsSpan(o, count).Fill(encoded[i++]);
                o += count;
            }
            else
            {
                output[o++] = b;
            }
        }

        return output;
    }

    private static int ReadUInt16(ReadOnlySpan<byte> data, int offset) => data[offset] | (data[offset + 1] << 8);
}
