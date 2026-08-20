using System.Buffers.Binary;
using System.Globalization;

namespace Invigoration.Core.Auth;

/// <summary>The product/public/private triple Blizzard's CD-key cipher embeds in a key, used to build the SID_AUTH_CHECK hash.</summary>
public readonly record struct DecodedCdKey(uint Product, uint Public, uint Private)
{
    /// <summary>The 20-byte X-SHA1 digest for this key, given this handshake's client/server tokens.</summary>
    public byte[] GetHash(uint clientToken, uint serverToken)
    {
        var buffer = new byte[24];
        var span = buffer.AsSpan();
        BinaryPrimitives.WriteUInt32LittleEndian(span[..4], clientToken);
        BinaryPrimitives.WriteUInt32LittleEndian(span[4..8], serverToken);
        BinaryPrimitives.WriteUInt32LittleEndian(span[8..12], Product);
        BinaryPrimitives.WriteUInt32LittleEndian(span[12..16], Public);
        BinaryPrimitives.WriteUInt32LittleEndian(span[16..20], 0); // reserved, always zero on the wire
        BinaryPrimitives.WriteUInt32LittleEndian(span[20..24], Private);
        return XSha1.Hash(buffer);
    }

    /// <summary>
    /// The full 36-byte CD-key block SID_AUTH_CHECK expects per key: Key
    /// Length(4) + Product(4) + Public(4) + reserved(4) + the 20-byte hash
    /// from <see cref="GetHash"/>. <paramref name="keyLength"/> is the
    /// length of the original CD-key string as typed (13 or 16).
    /// </summary>
    public byte[] GetAuthCheckBlock(int keyLength, uint clientToken, uint serverToken)
    {
        var block = new byte[36];
        var span = block.AsSpan();
        BinaryPrimitives.WriteUInt32LittleEndian(span[..4], (uint)keyLength);
        BinaryPrimitives.WriteUInt32LittleEndian(span[4..8], Product);
        BinaryPrimitives.WriteUInt32LittleEndian(span[8..12], Public);
        BinaryPrimitives.WriteUInt32LittleEndian(span[12..16], 0); // reserved
        GetHash(clientToken, serverToken).CopyTo(span[16..36]);
        return block;
    }
}

/// <summary>
/// Decodes classic 13-digit numeric (StarCraft/Diablo/Warcraft II) and modern
/// 16-character alphanumeric (Diablo II/Warcraft II:BNE) Battle.net CD-keys
/// into their embedded product/public/private values, so the SID_AUTH_CHECK
/// hash can be computed locally instead of round-tripping the raw key to
/// BNLS_CDKEY/BNLS_CDKEY_EX. Ported from the salt-substitution cipher in
/// Davnit/bncs.py's SCKeyDecoder/D2KeyDecoder (bncs/hashing/cdkeys.py) — the
/// original algorithm as reverse-engineered from Blizzard's client.
///
/// The 26-character Warcraft III/TFT key format uses a much more involved
/// bit-permutation cipher (a 30-round substitution-permutation network over
/// a 480-byte lookup table) that isn't ported here for lack of a verifiable
/// test vector; those two products still route their CD-key hash through
/// BNLS (see BotEngine.Bnls.cs) — WC3/TFT already depend on BNLS for their
/// NLS/SRP login challenge regardless, so this doesn't add a new dependency
/// for them.
/// </summary>
public static class CdKeyDecoder
{
    private const uint Salt = 0x13AC9741;

    private static readonly int[] ClassicAlpha = [6, 0, 2, 9, 3, 11, 1, 7, 5, 4, 10, 8];
    private static readonly int[] ModernAlpha = [5, 6, 0, 1, 2, 3, 4, 9, 10, 11, 12, 13, 14, 15, 7, 8];
    private const string ModernChars = "246789BCDEFGHJKMNPRTVWXZ";

    /// <summary>Decodes a key, dispatching on its length. Returns null if the key is malformed, uses an unsupported length, or fails its checksum.</summary>
    public static DecodedCdKey? Decode(string rawKey)
    {
        var key = rawKey.Trim().ToUpperInvariant();
        return key.Length switch
        {
            13 => DecodeClassic(key),
            16 => DecodeModern(key),
            _ => null,
        };
    }

    private static DecodedCdKey? DecodeClassic(string key)
    {
        if (!key.All(char.IsAsciiDigit))
        {
            return null;
        }

        var decoded = new int[12];
        var salt = Salt;

        for (var i = 11; i >= 0; i--)
        {
            var c = key[ClassicAlpha[i]];
            if (c <= 55) // '0'-'7'
            {
                decoded[i] = c ^ (int)(salt & 7);
                salt >>= 3;
            }
            else // '8'-'9'
            {
                decoded[i] = c ^ (i & 1);
            }
        }

        if (GetClassicCheckDigit(key) != key[12])
        {
            return null;
        }

        var value = new string(decoded.Select(v => (char)v).ToArray());
        if (!uint.TryParse(value[..2], out var product) ||
            !uint.TryParse(value[2..9], out var pub) ||
            !uint.TryParse(value[9..12], out var priv))
        {
            return null;
        }

        return new DecodedCdKey(product, pub, priv);
    }

    private static char GetClassicCheckDigit(string key)
    {
        var check = 3;
        for (var i = 0; i < 12; i++)
        {
            check += (key[i] - '0') ^ (check * 2);
        }

        return (char)('0' + (((check % 10) + 10) % 10));
    }

    private static DecodedCdKey? DecodeModern(string key)
    {
        var chars = key.ToCharArray();
        for (var i = 0; i < 15; i += 2)
        {
            var hi = ModernChars.IndexOf(chars[i]);
            var lo = ModernChars.IndexOf(chars[i + 1]);
            if (hi < 0 || lo < 0)
            {
                return null;
            }

            var n = (lo + (hi * 24)) & 0xFF;
            chars[i] = GetHexChar((n >> 4) & 0xF);
            chars[i + 1] = GetHexChar(n & 0xF);
        }

        var decoded = new int[16];
        var salt = Salt;

        for (var i = 15; i >= 0; i--)
        {
            var c = chars[ModernAlpha[i]];
            if (c <= 55) // '0'-'7'
            {
                decoded[i] = c ^ (int)(salt & 7);
                salt >>= 3;
            }
            else if (c < 65) // '8'-'9' (and the unused ':'-'@' range)
            {
                decoded[i] = c ^ (i & 1);
            }
            else // 'A'-'F' (already-decoded hex digit from the pass above)
            {
                decoded[i] = c;
            }
        }

        var hex = new string(decoded.Select(v => (char)v).ToArray());
        if (!uint.TryParse(hex[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var product) ||
            !uint.TryParse(hex[2..8], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var pub) ||
            !uint.TryParse(hex[8..16], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var priv))
        {
            return null;
        }

        return new DecodedCdKey(product, pub, priv);
    }

    private static char GetHexChar(int v)
    {
        v &= 0xF;
        return (char)(v < 10 ? v + '0' : (v - 10) + 'A');
    }
}
