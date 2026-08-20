using Invigoration.Sc2.Native;

namespace Invigoration.Sc2.Tests;

/// <summary>
/// Every hex vector here was copied directly from ncarrillo/superiority's
/// core/src/native/crypto.rs #[cfg(test)] block (fetched raw via `gh api`,
/// not relayed through a summary) — these are the exact assertions that
/// crate's own test suite makes.
/// </summary>
public class NativeCryptoTests
{
    [Fact]
    public void SessionProof_MatchesRecoveredHmacOrder()
    {
        var seed = Enumerable.Range(0, 64).Select(i => (byte)i).ToArray();
        var serverNonce = Enumerable.Range(16, 16).Select(i => (byte)i).ToArray();
        var clientNonce = Enumerable.Range(32, 16).Select(i => (byte)i).ToArray();

        var proof = NativeCrypto.BuildSessionProofWithNonce(seed, serverNonce, clientNonce);

        Assert.Equal(
            "680a8f22143bd8b198e251ebbbbd2404b3abf8e98726998e14bd1b20e1fb6927abd8d7432a79162e70ee6b88ac88f26cf705ffcd9d76ca02b131d70a95247c33",
            Convert.ToHexStringLower(proof.TransportKey));
        Assert.Equal(
            "d4926776eac9c008aa5702a99bf7ea01928db94b76afdb1f23a6156d47bbdc1f",
            Convert.ToHexStringLower(proof.Output[17..]));
        Assert.Equal(
            "ea649db50fff97f9691931772b87a3af161a6c5712ddce2cf6734249f92364cf",
            Convert.ToHexStringLower(proof.ExpectedServerProof));
    }

    [Fact]
    public void TransportSchedule_MatchesRecoveredVectors()
    {
        var key = Enumerable.Range(0, 64).Select(i => (byte)i).ToArray();
        var first = Enumerable.Range(0, 16).Select(i => (byte)i).ToArray();
        var second = Enumerable.Range(16, 16).Select(i => (byte)i).ToArray();

        var kdf = NativeCrypto.TransportKdf64(key, 2, first, second);

        Assert.Equal(
            "a1fe9de83671a75280e4ab71fcc689cc83d96d2690c073738a8e9282742267bafb37ee1073eae6501d81b7e769d43240ad345a3c85703122576259b90829ea45",
            Convert.ToHexStringLower(kdf));

        var (inbound, outbound) = NativeCrypto.DeriveTransportRc4Keys(key);

        Assert.Equal("5d179cef1b18de5f87e60f436a679e89ddea7ae7e92ec1a06c52df733ce124e2", Convert.ToHexStringLower(inbound));
        Assert.Equal("538690548cd9f6fa25c5204ef466f075757478decbfcc892be1e2c974c8637e3", Convert.ToHexStringLower(outbound));
    }

    [Fact]
    public void ThumbprintContext_NormalizesIpAddresses()
    {
        Assert.Equal(
            "00000000000000000000ffffc000020a",
            Convert.ToHexStringLower(NativeCrypto.ThumbprintContextForPeer("192.0.2.10")));
        Assert.Equal(
            "20010db8000000000000000000000001",
            Convert.ToHexStringLower(NativeCrypto.ThumbprintContextForPeer("2001:db8::1%en0")));
    }

    [Fact]
    public void VerifyThumbprint_AllZeroInput_Rejected()
    {
        // A real positive vector requires Blizzard's private key and doesn't exist in the
        // reference crate's own test suite either — this all-zero rejection is the one case
        // upstream's crypto.rs itself asserts (thumbprint_context_normalizes_ip_addresses).
        Assert.False(NativeCrypto.VerifyThumbprint(new byte[16], new byte[512]));
    }

    [Fact]
    public void VerifyThumbprint_WrongLengthInputs_Rejected()
    {
        // Also implicitly confirms the embedded 4096-bit modulus parsed without throwing:
        // a malformed constant would fail at static-init time before any assertion here ran.
        Assert.False(NativeCrypto.VerifyThumbprint(new byte[15], new byte[512]));
        Assert.False(NativeCrypto.VerifyThumbprint(new byte[16], new byte[511]));
    }
}

public class Rc4StateTests
{
    [Fact]
    public void Apply_MatchesClassicTestVector()
    {
        var expected = Convert.FromHexString("bbf316e8d940af0ad3");
        var plaintext = "Plaintext"u8.ToArray();

        var whole = new Rc4State("Key"u8.ToArray()).Apply(plaintext);

        Assert.Equal(expected, whole);
    }

    [Fact]
    public void Apply_PersistsStateAcrossFragmentedCalls()
    {
        var expected = Convert.FromHexString("bbf316e8d940af0ad3");
        var plaintext = "Plaintext"u8.ToArray();
        var cipher = new Rc4State("Key"u8.ToArray());

        var actual = cipher.Apply(plaintext.AsSpan(0, 2)).Concat(cipher.Apply(plaintext.AsSpan(2))).ToArray();

        Assert.Equal(expected, actual);
    }
}
