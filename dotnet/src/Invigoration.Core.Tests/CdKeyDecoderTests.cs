using Invigoration.Core.Auth;

namespace Invigoration.Core.Tests;

/// <summary>
/// CdKeyDecoder has no live-server test vector available (unlike XSha1), so
/// these tests cross-check it against ReferenceDecodeModern/EncodeClassic —
/// standalone re-implementations of the same cipher, typed independently
/// from CdKeyDecoder.cs rather than shared code. This catches indexing/logic
/// bugs (off-by-ones, wrong branch conditions, wrong operators), though it
/// can't rule out the same misunderstanding of the source algorithm being
/// made twice. Treat CD-key hashing as needing a live-login check before
/// fully trusting it.
/// </summary>
public class CdKeyDecoderTests
{
    [Fact]
    public void Decode_ClassicKey_RoundTripsThroughEncode()
    {
        var key = EncodeClassic(6, 1234567, 890);

        var decoded = CdKeyDecoder.Decode(key);

        Assert.NotNull(decoded);
        Assert.Equal(6u, decoded.Value.Product);
        Assert.Equal(1234567u, decoded.Value.Public);
        Assert.Equal(890u, decoded.Value.Private);
    }

    [Fact]
    public void Decode_ClassicKey_WrongCheckDigit_ReturnsNull()
    {
        var key = EncodeClassic(1, 654321, 42);
        var lastDigit = key[^1];
        var tampered = key[..^1] + (char)('0' + ((lastDigit - '0' + 1) % 10));

        var decoded = CdKeyDecoder.Decode(tampered);

        Assert.Null(decoded);
    }

    [Theory]
    [InlineData("2222222222222222")]
    [InlineData("XZWVTRPNMKJHGFED")]
    [InlineData("246897BCDEFGHJKM")]
    [InlineData("D2XPD2XPD2XPD2XP")]
    public void Decode_ModernKey_MatchesIndependentlyWrittenReferenceDecoder(string key)
    {
        var expected = ReferenceDecodeModern(key);
        var actual = CdKeyDecoder.Decode(key);

        Assert.Equal(expected, actual);
    }

    [Fact]
    public void Decode_UnsupportedLength_ReturnsNull()
    {
        Assert.Null(CdKeyDecoder.Decode("TOOSHORT"));
    }

    [Fact]
    public void GetHash_ProducesTwentyByteDigest()
    {
        var decoded = new DecodedCdKey(6, 1234567, 890);

        var hash = decoded.GetHash(0x11223344, 0x55667788);

        Assert.Equal(20, hash.Length);
    }

    [Fact]
    public void GetAuthCheckBlock_ProducesThirtySixByteBlockWithHeaderAndHash()
    {
        var decoded = new DecodedCdKey(6, 1234567, 890);

        var block = decoded.GetAuthCheckBlock(13, 0x11223344, 0x55667788);

        Assert.Equal(36, block.Length);
        Assert.Equal(13u, BitConverter.ToUInt32(block, 0));
        Assert.Equal(6u, BitConverter.ToUInt32(block, 4));
        Assert.Equal(1234567u, BitConverter.ToUInt32(block, 8));
        Assert.Equal(0u, BitConverter.ToUInt32(block, 12));
        Assert.Equal(decoded.GetHash(0x11223344, 0x55667788), block[16..36]);
    }

    // --- Standalone re-implementation of the encode side, typed independently from CdKeyDecoder.cs ---

    private static readonly int[] ClassicAlpha = [6, 0, 2, 9, 3, 11, 1, 7, 5, 4, 10, 8];
    private static readonly int[] ModernAlpha = [5, 6, 0, 1, 2, 3, 4, 9, 10, 11, 12, 13, 14, 15, 7, 8];
    private const string ModernChars = "246789BCDEFGHJKMNPRTVWXZ";
    private const uint Salt = 0x13AC9741;

    private static string EncodeClassic(int product, int publicValue, int privateValue)
    {
        var plain = $"{product:D2}{publicValue:D7}{privateValue:D3}";
        var encoded = new char[12];
        var salt = Salt;

        for (var i = 11; i >= 0; i--)
        {
            int c = plain[i];
            var target = ClassicAlpha[i];
            encoded[target] = c <= 55
                ? (char)(c ^ (int)(salt & 7))
                : (char)(c ^ (i & 1));
            if (c <= 55)
            {
                salt >>= 3;
            }
        }

        var check = 3;
        for (var i = 0; i < 12; i++)
        {
            check += (encoded[i] - '0') ^ (check * 2);
        }

        return new string(encoded) + (char)('0' + (((check % 10) + 10) % 10));
    }

    private static DecodedCdKey? ReferenceDecodeModern(string rawKey)
    {
        var key = rawKey.Trim().ToUpperInvariant();
        if (key.Length != 16)
        {
            return null;
        }

        var chars = key.ToCharArray();
        for (var i = 0; i < 16; i += 2)
        {
            var hiIndex = ModernChars.IndexOf(chars[i]);
            var loIndex = ModernChars.IndexOf(chars[i + 1]);
            if (hiIndex < 0 || loIndex < 0)
            {
                return null;
            }

            var n = (loIndex + (hiIndex * 24)) & 0xFF;
            chars[i] = ToHexChar((n >> 4) & 0xF);
            chars[i + 1] = ToHexChar(n & 0xF);
        }

        var plain = new char[16];
        var salt = Salt;
        for (var i = 15; i >= 0; i--)
        {
            int c = chars[ModernAlpha[i]];
            if (c <= 55)
            {
                plain[i] = (char)(c ^ (int)(salt & 7));
                salt >>= 3;
            }
            else if (c < 65)
            {
                plain[i] = (char)(c ^ (i & 1));
            }
            else
            {
                plain[i] = (char)c;
            }
        }

        var hex = new string(plain);
        var product = Convert.ToUInt32(hex[..2], 16);
        var pub = Convert.ToUInt32(hex[2..8], 16);
        var priv = Convert.ToUInt32(hex[8..16], 16);
        return new DecodedCdKey(product, pub, priv);
    }

    private static char ToHexChar(int v) => (char)(v < 10 ? v + '0' : (v - 10) + 'A');
}
