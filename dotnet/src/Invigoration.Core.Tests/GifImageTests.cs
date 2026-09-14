using Invigoration.Core.Imaging;

namespace Invigoration.Core.Tests;

/// <summary>Builds GIFs for tests: frames of palette indices, LZW-coded without compression (a clear code before the table would widen the codes), which any decoder must still read.</summary>
internal static class TestGif
{
    public static byte[] Build(int width, int height, params byte[][] frames) =>
        Build(width, height, [.. frames.Select(f => (0, 0, width, height, f))], interlaced: false);

    public static byte[] Build(int width, int height, IReadOnlyList<(int Left, int Top, int Width, int Height, byte[] Pixels)> frames, bool interlaced)
    {
        var gif = new List<byte>();
        gif.AddRange("GIF89a"u8.ToArray());
        Word(gif, width);
        Word(gif, height);
        gif.AddRange([0xF7, 0, 0]); // 256-colour global table follows
        for (var i = 0; i < 256; i++)
        {
            gif.AddRange([(byte)i, (byte)i, (byte)i]);
        }

        foreach (var (left, top, w, h, pixels) in frames)
        {
            gif.AddRange([0x21, 0xF9, 4, 0x09, 10, 0, 0, 0]); // disposal 2, transparent index 0
            gif.Add(0x2C);
            Word(gif, left);
            Word(gif, top);
            Word(gif, w);
            Word(gif, h);
            gif.Add(interlaced ? (byte)0x40 : (byte)0);
            gif.Add(8);

            var bits = new BitWriter();
            for (var i = 0; i < pixels.Length; i++)
            {
                if (i % 254 == 0)
                {
                    bits.Write(256, 9);
                }

                bits.Write(pixels[i], 9);
            }

            bits.Write(257, 9);
            var data = bits.ToArray();
            for (var i = 0; i < data.Length; i += 255)
            {
                var n = Math.Min(255, data.Length - i);
                gif.Add((byte)n);
                gif.AddRange(data.AsSpan(i, n).ToArray());
            }

            gif.Add(0);
        }

        gif.Add(0x3B);
        return [.. gif];
    }

    private static void Word(List<byte> bytes, int value)
    {
        bytes.Add((byte)value);
        bytes.Add((byte)(value >> 8));
    }

    private sealed class BitWriter
    {
        private readonly List<byte> _bytes = [];
        private int _buffer;
        private int _count;

        public void Write(int code, int size)
        {
            _buffer |= code << _count;
            _count += size;
            while (_count >= 8)
            {
                _bytes.Add((byte)_buffer);
                _buffer >>= 8;
                _count -= 8;
            }
        }

        public byte[] ToArray() => _count > 0 ? [.. _bytes, (byte)_buffer] : [.. _bytes];
    }
}

public class GifImageTests
{
    // A 40×30, two-frame GIF saved by Pillow with real LZW compression (so the code width grows),
    // where frame f's pixel (x, y) is (x * 7 + y * 3 + f * 5) % 23, or 0 where (x + y + f) % 5 == 0.
    private const string PillowGif =
        "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";

    private static byte Expected(int x, int y, int f) => (x + y + f) % 5 == 0 ? (byte)0 : (byte)((x * 7 + y * 3 + f * 5) % 23);

    [Fact]
    public void Decode_ReadsCompressedFramesFromARealEncoder()
    {
        var gif = GifImage.Decode(Convert.FromBase64String(PillowGif));

        Assert.Equal(40, gif.Width);
        Assert.Equal(30, gif.Height);
        Assert.Equal(2, gif.Frames.Count);
        for (var f = 0; f < 2; f++)
        {
            for (var y = 0; y < 30; y++)
            {
                for (var x = 0; x < 40; x++)
                {
                    Assert.Equal(Expected(x, y, f), gif.Frames[f][y * 40 + x]);
                }
            }
        }
    }

    [Fact]
    public void Decode_ReadsUncompressedCodesAcrossClearCodes()
    {
        var pixels = Enumerable.Range(0, 600).Select(i => (byte)(i % 251)).ToArray();

        var gif = GifImage.Decode(TestGif.Build(30, 20, pixels));

        Assert.Equal(pixels, gif.Frames[0]);
    }

    [Fact]
    public void Decode_PlacesASmallerFrameOnATransparentCanvas()
    {
        var gif = GifImage.Decode(TestGif.Build(4, 3, [(1, 1, 2, 2, new byte[] { 5, 6, 7, 8 })], interlaced: false));

        Assert.Equal(new byte[] { 0, 0, 0, 0, 0, 5, 6, 0, 0, 7, 8, 0 }, gif.Frames[0]);
    }

    [Fact]
    public void Decode_RejectsWhatItCantRead()
    {
        Assert.Throws<FormatException>(() => GifImage.Decode("PNG89a..."u8));
        Assert.Throws<FormatException>(() => GifImage.Decode(TestGif.Build(2, 2, [(0, 0, 2, 2, new byte[] { 1, 2, 3, 4 })], interlaced: true)));
        Assert.Throws<FormatException>(() => GifImage.Decode(TestGif.Build(2, 2, [(1, 1, 2, 2, new byte[] { 1, 2, 3, 4 })], interlaced: false)));

        var whole = TestGif.Build(2, 2, new byte[] { 1, 2, 3, 4 });
        Assert.Throws<FormatException>(() => GifImage.Decode(whole.AsSpan(0, whole.Length / 2)));
    }

    [Fact]
    public void Decode_StopsAtTheFrameLimit()
    {
        var gif = TestGif.Build(1, 1, [.. Enumerable.Repeat(new byte[] { 1 }, 5)]);

        Assert.Equal(5, GifImage.Decode(gif, maxFrames: 5).Frames.Count);
        Assert.Throws<FormatException>(() => GifImage.Decode(gif, maxFrames: 4));
    }
}
