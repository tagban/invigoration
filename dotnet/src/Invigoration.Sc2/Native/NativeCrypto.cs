using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;

namespace Invigoration.Sc2.Native;

/// <summary>
/// HMAC-SHA256-based key derivation for SC2's native ("Sunken") transport:
/// the session proof exchanged during resume, the RC4 keys derived from it,
/// and RSA thumbprint verification of the server's identity. Ported from
/// core/src/native/crypto.rs.
/// </summary>
public static class NativeCrypto
{
    // Same 512-byte little-endian RSA-4096 modulus as core/src/native/crypto.rs's THUMBPRINT_MODULUS_LE.
    private const string ThumbprintModulusLittleEndianHex =
        "c95435d383bf6622543f9e30f301c545942854e28cef54df37542c079534d25d84ff3fe6d0fa5ed2cfe62fd44cfb4c" +
        "d6911c0d90328a2adc92c3c5e7774e7b50ac1e38da03c77732899c17cf7bd3872b7ad82cd198fee536545a593da9635" +
        "a293dc322a9a55fdb9269b0ff7b4b8efcca69317d2dfceae629588f4653295a3ee68ca8a013b2151d60be840a76f4c9" +
        "b54e4827d1c39d313012a2ecf1e4826e9c78ba46a24856107821b8f7d150676bf78d0ec54ea49840075a4a92b152272" +
        "c21d86b2e67fb36567ffdbb9af551bfa743e72409b8a009f33c5bf100c135d34f9aef33f4f1de67ab28d7a750f4b103" +
        "f788482668d626d6fc3d1b30f1a2800c17feeb3016c035fe406d6378bb66b61d008bc6213bc855ff0ff32330b61ed41" +
        "73604b7f62ebc8c09c4c7c0d259fde67c1044b156d5f63e42c33b2239d733258882f07ef1ea79e43eee1365dfd51bef" +
        "79a68a815c31146febc348459870aa1cfdb7f2b6fcd173cac05967b5950337af7e213d34eb1217b59c8dd6aeaa88ceb" +
        "800bc277f97a734281e68ca4836bf2d6e1b8f7bec3763df1c3ec3b30f39731167cac86f75845294717c6c663e48dca6" +
        "87bfbcc39786158e44cdd3b144df84933d502041f3abe588f96d342000d650be5afa939985b1272784d1c7b1fea3fa4" +
        "9f1fc028ebcdacc4ded8749fd40cd50458ee1c30c4d2c1b3c9bcfb4b3696a2bf40dde83afde";

    private static readonly byte[] ThumbprintDomain = "Thumbprint.IPv6"u8.ToArray();

    private static readonly RSA ThumbprintPublicKey = CreateThumbprintPublicKey();

    private static RSA CreateThumbprintPublicKey()
    {
        var modulusLittleEndian = Convert.FromHexString(ThumbprintModulusLittleEndianHex);
        Array.Reverse(modulusLittleEndian);
        var rsa = RSA.Create();
        rsa.ImportParameters(new RSAParameters { Modulus = modulusLittleEndian, Exponent = [0x01, 0x00, 0x01] });
        return rsa;
    }

    /// <summary>
    /// Verifies the server's 4096-bit RSA PKCS#1 v1.5 SHA-512 signature over
    /// <paramref name="peerAddressContext"/> (the 16-byte IPv6-mapped peer
    /// address from <see cref="ThumbprintContextForPeer"/>) concatenated
    /// with the ASCII domain "Thumbprint.IPv6". The wire signature is
    /// little-endian; it is byte-reversed before verification.
    /// </summary>
    public static bool VerifyThumbprint(byte[] peerAddressContext, byte[] signatureLittleEndian)
    {
        if (peerAddressContext.Length != 16 || signatureLittleEndian.Length != 512)
        {
            return false;
        }

        var digest = SHA512.HashData([.. peerAddressContext, .. ThumbprintDomain]);
        var signature = (byte[])signatureLittleEndian.Clone();
        Array.Reverse(signature);
        return ThumbprintPublicKey.VerifyHash(digest, signature, HashAlgorithmName.SHA512, RSASignaturePadding.Pkcs1);
    }
    private static readonly byte[] InboundRc4Label =
    [
        0x68, 0xe0, 0xc7, 0x2e, 0xdd, 0xd6, 0xd2, 0xf3, 0x1e, 0x5a, 0xb1, 0x55, 0xb1, 0x8b, 0x63, 0x1e,
    ];

    private static readonly byte[] OutboundRc4Label =
    [
        0xde, 0xa9, 0x65, 0xae, 0x54, 0x3a, 0x1e, 0x93, 0x9e, 0x69, 0x0c, 0xaa, 0x68, 0xde, 0x78, 0x39,
    ];

    /// <summary>Two-block HMAC schedule used both to negotiate the next transport context and to derive a fresh protected secret (domain 0 and 2 respectively).</summary>
    public static byte[] TransportKdf64(byte[] key, byte domain, byte[] firstContext, byte[] secondContext)
    {
        RequireLength(key, 64, nameof(key));
        RequireLength(firstContext, 16, nameof(firstContext));
        RequireLength(secondContext, 16, nameof(secondContext));

        var first = HmacSha256(key, [domain], firstContext, secondContext);
        var second = HmacSha256(key, [domain], secondContext, firstContext, [domain]);

        var output = new byte[64];
        first.CopyTo(output, 0);
        second.CopyTo(output, 32);
        return output;
    }

    public static byte[] DeriveSessionAuthKey(byte[] sessionSeed, byte[] clientNonce, byte[] serverNonce)
    {
        RequireLength(sessionSeed, 64, nameof(sessionSeed));
        RequireLength(clientNonce, 16, nameof(clientNonce));
        RequireLength(serverNonce, 16, nameof(serverNonce));

        var first = HmacSha256(sessionSeed, [0x00], clientNonce, serverNonce);
        var second = HmacSha256(sessionSeed, [0x01], serverNonce, clientNonce);

        var output = new byte[64];
        first.CopyTo(output, 0);
        second.CopyTo(output, 32);
        return output;
    }

    public static (byte[] Inbound, byte[] Outbound) DeriveTransportRc4Keys(byte[] protectedSecret)
    {
        RequireLength(protectedSecret, 64, nameof(protectedSecret));
        return (
            HmacSha256(protectedSecret, InboundRc4Label),
            HmacSha256(protectedSecret, OutboundRc4Label));
    }

    public sealed record SessionProof(
        byte[] Output,
        byte[] ClientNonce,
        byte[] ServerNonce,
        byte[] TransportKey,
        byte[] ExpectedServerProof);

    /// <summary>Builds the 49-byte session proof sent to the server: [phase=1][client_nonce(16)][client_proof(32)].</summary>
    public static SessionProof BuildSessionProofWithNonce(byte[] sessionSeed, byte[] serverNonce, byte[] clientNonce)
    {
        RequireLength(sessionSeed, 64, nameof(sessionSeed));
        RequireLength(serverNonce, 16, nameof(serverNonce));
        RequireLength(clientNonce, 16, nameof(clientNonce));

        var transportKey = DeriveSessionAuthKey(sessionSeed, clientNonce, serverNonce);
        var clientProof = HmacSha256(transportKey, [0x00], clientNonce, serverNonce);
        var expectedServerProof = HmacSha256(transportKey, [0x01], serverNonce, clientNonce);

        var output = new byte[49];
        output[0] = 1;
        clientNonce.CopyTo(output, 1);
        clientProof.CopyTo(output, 17);

        return new SessionProof(output, clientNonce, serverNonce, transportKey, expectedServerProof);
    }

    public static SessionProof BuildSessionProof(byte[] sessionSeed, byte[] serverNonce) =>
        BuildSessionProofWithNonce(sessionSeed, serverNonce, RandomNumberGenerator.GetBytes(16));

    public sealed record TransportHandshake(
        byte[] ClientContext,
        byte[] ServerContext,
        byte[] Response,
        byte[] ProtectedSecret);

    /// <summary>
    /// Periodic transport re-key (Conn/11 regulator maintenance): derives a
    /// fresh protected secret from the current transport key plus a new
    /// client/server context pair, returning the 48-byte response to send
    /// and the secret to feed into <see cref="DeriveTransportRc4Keys"/> next.
    /// </summary>
    public static TransportHandshake BuildTransportHandshakeWithNonce(byte[] transportKey, byte[] serverContext, byte[] clientContext)
    {
        RequireLength(transportKey, 64, nameof(transportKey));
        RequireLength(serverContext, 16, nameof(serverContext));
        RequireLength(clientContext, 16, nameof(clientContext));

        var negotiation = TransportKdf64(transportKey, 0, clientContext, serverContext);
        var nextContext = negotiation[32..48];
        var protectedSecret = TransportKdf64(transportKey, 2, nextContext, serverContext);

        var response = new byte[48];
        nextContext.CopyTo(response, 0);
        negotiation.AsSpan(0, 32).CopyTo(response.AsSpan(16));

        return new TransportHandshake(clientContext, serverContext, response, protectedSecret);
    }

    public static TransportHandshake BuildTransportHandshake(byte[] transportKey, byte[] serverContext) =>
        BuildTransportHandshakeWithNonce(transportKey, serverContext, RandomNumberGenerator.GetBytes(16));

    /// <summary>The 16-byte IPv6-normalized peer address used as thumbprint verification context (IPv4 peers are mapped to ::ffff:a.b.c.d first).</summary>
    public static byte[] ThumbprintContextForPeer(string address)
    {
        var host = address.Split('%')[0];
        var ip = IPAddress.Parse(host);
        return (ip.AddressFamily == AddressFamily.InterNetwork ? ip.MapToIPv6() : ip).GetAddressBytes();
    }

    private static byte[] HmacSha256(byte[] key, params byte[][] parts)
    {
        using var hmac = new HMACSHA256(key);
        foreach (var part in parts)
        {
            hmac.TransformBlock(part, 0, part.Length, null, 0);
        }

        hmac.TransformFinalBlock([], 0, 0);
        return hmac.Hash!;
    }

    private static void RequireLength(byte[] value, int expected, string name)
    {
        if (value.Length != expected)
        {
            throw new ArgumentException($"{name} must be {expected} bytes, was {value.Length}.", name);
        }
    }
}
