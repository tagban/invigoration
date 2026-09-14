namespace Invigoration.Core.Imaging;

/// <summary>
/// Decodes an animated GIF to palette indices — no colours, since the Diablo II character pack
/// (<see cref="StatString.D2CharacterPack"/>) remaps indices through its own tint tables before
/// they become colours. Each frame is a full canvas: a frame drawn into a smaller rectangle is
/// placed on a canvas of the transparent index (every pack GIF clears between frames), and
/// interlaced frames are rejected rather than half-supported. Input is data from the network,
/// so every size is checked before anything is allocated from it.
/// </summary>
public sealed class GifImage
{
    public const int MaxDimension = 4096;

    private GifImage(int width, int height, IReadOnlyList<byte[]> frames)
    {
        Width = width;
        Height = height;
        Frames = frames;
    }

    public int Width { get; }

    public int Height { get; }

    /// <summary>One Width × Height array of palette indices per frame, rows top to bottom.</summary>
    public IReadOnlyList<byte[]> Frames { get; }

    /// <param name="data">The GIF file.</param>
    /// <param name="maxFrames">Stop with an error past this many frames.</param>
    /// <exception cref="FormatException">Not a GIF this decoder can read.</exception>
    public static GifImage Decode(ReadOnlySpan<byte> data, int maxFrames = 256)
    {
        var reader = new Reader(data);
        var signature = reader.Bytes(6);
        if (!signature.SequenceEqual("GIF89a"u8) && !signature.SequenceEqual("GIF87a"u8))
        {
            throw new FormatException("Not a GIF.");
        }

        var width = reader.Word();
        var height = reader.Word();
        if (width is 0 or > MaxDimension || height is 0 or > MaxDimension)
        {
            throw new FormatException($"Unusable GIF size {width}×{height}.");
        }

        var flags = reader.Byte();
        reader.Skip(2); // background colour, aspect ratio
        if ((flags & 0x80) != 0)
        {
            reader.Skip(3 * (2 << (flags & 7)));
        }

        var frames = new List<byte[]>();
        var transparentIndex = 0;
        while (true)
        {
            switch (reader.Byte())
            {
                case 0x21: // extension
                    var label = reader.Byte();
                    if (label == 0xF9 && reader.Peek() >= 4)
                    {
                        reader.Skip(1); // block size
                        var gceFlags = reader.Byte();
                        reader.Skip(2); // delay
                        var index = reader.Byte();
                        transparentIndex = (gceFlags & 1) != 0 ? index : 0;
                    }

                    SkipSubBlocks(ref reader);
                    break;

                case 0x2C: // image
                    if (frames.Count >= maxFrames)
                    {
                        throw new FormatException("Too many GIF frames.");
                    }

                    frames.Add(DecodeFrame(ref reader, width, height, (byte)transparentIndex));
                    break;

                case 0x3B: // trailer
                    if (frames.Count == 0)
                    {
                        throw new FormatException("A GIF with no frames.");
                    }

                    return new GifImage(width, height, frames);

                default:
                    throw new FormatException("Corrupt GIF block.");
            }
        }
    }

    private static byte[] DecodeFrame(ref Reader reader, int width, int height, byte transparentIndex)
    {
        var left = reader.Word();
        var top = reader.Word();
        var frameWidth = reader.Word();
        var frameHeight = reader.Word();
        var flags = reader.Byte();
        if ((flags & 0x80) != 0)
        {
            reader.Skip(3 * (2 << (flags & 7)));
        }

        if ((flags & 0x40) != 0)
        {
            throw new FormatException("Interlaced GIF frames aren't supported.");
        }

        if (left + frameWidth > width || top + frameHeight > height)
        {
            throw new FormatException("A GIF frame outside its canvas.");
        }

        var minCodeSize = reader.Byte();
        if (minCodeSize is < 2 or > 11)
        {
            throw new FormatException("Corrupt GIF image data.");
        }

        // Gather the frame's LZW sub-blocks into one run.
        var compressed = new List<byte>();
        while (reader.Byte() is var length and > 0)
        {
            compressed.AddRange(reader.Bytes(length));
        }

        var pixels = new byte[frameWidth * frameHeight];
        var count = Lzw(compressed, minCodeSize, pixels);
        if (count < pixels.Length)
        {
            // Some encoders stop a little short; the rest stays transparent.
            Array.Fill(pixels, transparentIndex, count, pixels.Length - count);
        }

        if (frameWidth == width && frameHeight == height)
        {
            return pixels;
        }

        var canvas = new byte[width * height];
        Array.Fill(canvas, transparentIndex);
        for (var y = 0; y < frameHeight; y++)
        {
            Array.Copy(pixels, y * frameWidth, canvas, (top + y) * width + left, frameWidth);
        }

        return canvas;
    }

    /// <summary>GIF's variable-width LZW. Returns how many pixels it produced.</summary>
    private static int Lzw(List<byte> input, int minCodeSize, byte[] output)
    {
        const int maxCodes = 4096;
        var prefix = new short[maxCodes];
        var suffix = new byte[maxCodes];
        var stack = new byte[maxCodes + 1];

        var clear = 1 << minCodeSize;
        var end = clear + 1;
        var codeSize = minCodeSize + 1;
        var next = clear + 2;
        var previous = -1;
        byte first = 0;

        for (var i = 0; i < clear; i++)
        {
            prefix[i] = -1;
            suffix[i] = (byte)i;
        }

        var written = 0;
        var bits = 0;
        var bitCount = 0;
        var position = 0;
        while (written < output.Length)
        {
            while (bitCount < codeSize)
            {
                if (position >= input.Count)
                {
                    return written;
                }

                bits |= input[position++] << bitCount;
                bitCount += 8;
            }

            var code = bits & ((1 << codeSize) - 1);
            bits >>= codeSize;
            bitCount -= codeSize;

            if (code == clear)
            {
                codeSize = minCodeSize + 1;
                next = clear + 2;
                previous = -1;
                continue;
            }

            if (code == end)
            {
                return written;
            }

            if (previous == -1)
            {
                if (code >= clear)
                {
                    throw new FormatException("Corrupt GIF image data.");
                }

                output[written++] = suffix[code];
                previous = code;
                first = suffix[code];
                continue;
            }

            var current = code;
            var depth = 0;
            if (code >= next)
            {
                if (code > next)
                {
                    throw new FormatException("Corrupt GIF image data.");
                }

                stack[depth++] = first;
                current = previous;
            }

            while (current >= clear)
            {
                stack[depth++] = suffix[current];
                current = prefix[current];
                if (depth >= maxCodes)
                {
                    throw new FormatException("Corrupt GIF image data.");
                }
            }

            first = suffix[current];
            stack[depth++] = first;

            while (depth > 0 && written < output.Length)
            {
                output[written++] = stack[--depth];
            }

            if (next < maxCodes)
            {
                prefix[next] = (short)previous;
                suffix[next] = first;
                next++;
                if (next == 1 << codeSize && codeSize < 12)
                {
                    codeSize++;
                }
            }

            previous = code;
        }

        return written;
    }

    private static void SkipSubBlocks(ref Reader reader)
    {
        while (reader.Byte() is var length and > 0)
        {
            reader.Skip(length);
        }
    }

    private ref struct Reader(ReadOnlySpan<byte> data)
    {
        private readonly ReadOnlySpan<byte> _data = data;
        private int _position;

        public byte Byte() => _position < _data.Length ? _data[_position++] : throw Truncated();

        public byte Peek() => _position < _data.Length ? _data[_position] : throw Truncated();

        public int Word() => Byte() | (Byte() << 8);

        public ReadOnlySpan<byte> Bytes(int count)
        {
            if (count > _data.Length - _position)
            {
                throw Truncated();
            }

            var span = _data.Slice(_position, count);
            _position += count;
            return span;
        }

        public void Skip(int count) => Bytes(count);

        private static FormatException Truncated() => new("The GIF ends early.");
    }
}
