using System.Text;
using Org.BouncyCastle.Crypto.Engines;

namespace Invigoration.Core.Music.Pandora;

/// <summary>
/// Pandora's JSON-RPC API encrypts every request body (except auth.partnerLogin itself) with
/// Blowfish in ECB mode, PKCS#7-padded to Blowfish's 8-byte block size, then hex-encoded — ported
/// from pydora's <c>BlowfishCryptor</c>/<c>Encryptor</c> rather than reimplemented from the docs
/// page alone (which had gaps). BouncyCastle's <see cref="BlowfishEngine"/> only exposes raw
/// single-block processing, not a ready-made ECB mode, so this loops over 8-byte blocks itself —
/// that's exactly what ECB mode is (no chaining between blocks), so there's nothing missing here.
/// </summary>
public static class PandoraCrypto
{
    private const int BlockSize = 8;

    public static string Encrypt(string key, string plaintext)
    {
        var engine = MakeEngine(key, forEncryption: true);
        var padded = Pad(Encoding.UTF8.GetBytes(plaintext));
        var output = new byte[padded.Length];
        for (var offset = 0; offset < padded.Length; offset += BlockSize)
        {
            engine.ProcessBlock(padded, offset, output, offset);
        }

        return ToHex(output);
    }

    /// <summary>General-purpose decrypt, used for actual response bodies — unpads the result.</summary>
    public static string Decrypt(string key, string hexCiphertext)
    {
        var raw = DecryptRaw(key, hexCiphertext);
        return Encoding.UTF8.GetString(Unpad(raw));
    }

    /// <summary>
    /// auth.partnerLogin's response embeds <c>syncTime</c> Blowfish-encrypted (with the partner's
    /// decrypt key) but NOT PKCS#7-padded the normal way — pydora reads it as the raw decrypted
    /// bytes with the first 4 and last 2 bytes trimmed off, not as padded UTF-8 text. Kept
    /// separate from <see cref="Decrypt"/> rather than folding a special case into it, since this
    /// shape is unique to that one response field.
    /// </summary>
    public static byte[] DecryptRaw(string key, string hexCiphertext)
    {
        var ciphertext = FromHex(hexCiphertext);
        var engine = MakeEngine(key, forEncryption: false);
        var output = new byte[ciphertext.Length];
        for (var offset = 0; offset < ciphertext.Length; offset += BlockSize)
        {
            engine.ProcessBlock(ciphertext, offset, output, offset);
        }

        return output;
    }

    /// <summary>
    /// auth.partnerLogin's syncTime, once Blowfish-decrypted, is the ASCII decimal digits of the
    /// unix timestamp wrapped in 4 leading + 2 trailing junk bytes — confirmed directly against
    /// pydora's own test fixture (decrypting "1507411159" ASCII-as-bytes slices to "4111", exactly
    /// Python's <c>[4:-2]</c> string slicing), not a raw binary integer as the docs page alone
    /// would suggest.
    /// </summary>
    public static long ParseSyncTimeDigits(byte[] decryptedRaw)
    {
        var digits = Encoding.ASCII.GetString(decryptedRaw, 4, decryptedRaw.Length - 6);
        return long.Parse(digits);
    }

    private static BlowfishEngine MakeEngine(string key, bool forEncryption)
    {
        var engine = new BlowfishEngine();
        engine.Init(forEncryption, new Org.BouncyCastle.Crypto.Parameters.KeyParameter(Encoding.UTF8.GetBytes(key)));
        return engine;
    }

    /// <summary>PKCS#7-style: pad to the next 8-byte boundary with pad_size bytes each equal to pad_size itself — e.g. 6 bytes of data gets 2 bytes of value 0x02 (confirmed against pydora's own test: "123456" encrypts to "123456\x02\x02" before the Blowfish step). Public because it's a pure, independently-verifiable function worth testing directly rather than only through a full encrypt/decrypt round trip.</summary>
    public static byte[] Pad(byte[] data)
    {
        var padSize = BlockSize - (data.Length % BlockSize);
        var padded = new byte[data.Length + padSize];
        Array.Copy(data, padded, data.Length);
        for (var i = data.Length; i < padded.Length; i++)
        {
            padded[i] = (byte)padSize;
        }

        return padded;
    }

    public static byte[] Unpad(byte[] data)
    {
        if (data.Length == 0)
        {
            return data;
        }

        var padSize = data[^1];
        if (padSize == 0 || padSize > BlockSize || padSize > data.Length)
        {
            return data;
        }

        return data[..^padSize];
    }

    private static string ToHex(byte[] data) => Convert.ToHexStringLower(data);

    private static byte[] FromHex(string hex) => Convert.FromHexString(hex);
}
