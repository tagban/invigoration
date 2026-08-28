using System.Text;
using Invigoration.Core.Music.Pandora;

namespace Invigoration.Core.Tests;

/// <summary>
/// Verifies PandoraCrypto's own glue logic (padding math, hex round-trip, syncTime slicing)
/// against pydora's real test fixtures (github.com/mcrute/pydora/blob/master/tests/test_pandora/
/// test_transport.py, fetched directly rather than guessed) — actual Blowfish correctness is
/// BouncyCastle's problem, not this codebase's, so these tests target the parts we actually wrote.
/// </summary>
public class PandoraCryptoTests
{
    [Fact]
    public void Pad_SixByteInput_AddsTwoBytesOfValueTwo()
    {
        // pydora: data = "123456" -> self.assertEqual(b"123456\x02\x02", cryptor.encrypt(data)) (pre-Blowfish).
        var padded = PandoraCrypto.Pad(Encoding.ASCII.GetBytes("123456"));

        Assert.Equal(Encoding.ASCII.GetBytes("123456\x02\x02"), padded);
    }

    [Fact]
    public void Pad_ExactBlockMultiple_AddsFullBlockOfPadding()
    {
        var padded = PandoraCrypto.Pad(Encoding.ASCII.GetBytes("12345678"));

        Assert.Equal(16, padded.Length);
        Assert.All(padded[8..], b => Assert.Equal(8, b));
    }

    [Fact]
    public void Unpad_StripsTrailingPaddingBytes()
    {
        // pydora: data = b"123456\x02\x02" -> self.assertEqual(b"123456", cryptor.decrypt(data)).
        var unpadded = PandoraCrypto.Unpad(Encoding.ASCII.GetBytes("123456\x02\x02"));

        Assert.Equal(Encoding.ASCII.GetBytes("123456"), unpadded);
    }

    [Fact]
    public void ParseSyncTimeDigits_MatchesPydoraFixture()
    {
        // pydora: ENCODED_TIME = "31353037343131313539" (hex) decodes to ASCII "1507411159";
        // decrypt_sync_time slices [4:-2] -> "4111" -> EXPECTED_TIME = 4111.
        var decrypted = Encoding.ASCII.GetBytes("1507411159");

        var result = PandoraCrypto.ParseSyncTimeDigits(decrypted);

        Assert.Equal(4111, result);
    }

    [Fact]
    public void EncryptThenDecrypt_RoundTripsArbitraryJson()
    {
        const string key = "6#26FRL$ZWD";
        const string plaintext = """{"foo":"bar"}""";

        var ciphertext = PandoraCrypto.Encrypt(key, plaintext);
        var decrypted = PandoraCrypto.Decrypt(key, ciphertext);

        Assert.Equal(plaintext, decrypted);
    }

    [Fact]
    public void EncryptThenDecrypt_RoundTripsExactBlockMultiple()
    {
        const string key = "R=U!LH$O2B#";
        const string plaintext = "exactly16chars!!";
        Assert.Equal(16, plaintext.Length);

        var ciphertext = PandoraCrypto.Encrypt(key, plaintext);
        var decrypted = PandoraCrypto.Decrypt(key, ciphertext);

        Assert.Equal(plaintext, decrypted);
    }

    [Fact]
    public void Encrypt_ProducesLowercaseHex()
    {
        var ciphertext = PandoraCrypto.Encrypt("somekey", "data");

        Assert.Matches("^[0-9a-f]+$", ciphertext);
    }
}
