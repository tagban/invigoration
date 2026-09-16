using System.Buffers.Binary;
using System.Net;
using System.Security.Cryptography;
using System.Text;

namespace Invigoration.Core.Hotline;

/// <summary>
/// HOPE — the Hotline One-time Password Extension, the modern ("3.x"-era) login used by HotSocket,
/// Janus, Hermes and friends. Classic Hotline sends the password bitwise-inverted, which is not
/// encryption at all: anyone watching the connection reads it. HOPE replaces that with a MAC over
/// a server-issued challenge, so the password itself never crosses the wire.
///
/// It rides inside the ordinary Login (107) transaction, so nothing new is framed:
///
///   1. The client sends a Login whose login field is a single 0x00 byte, plus the MAC algorithms
///      it supports and its own name. The null login is the signal — a server that doesn't know
///      HOPE just sees a failed login, and a client that omits the algorithm list gets the classic
///      path.
///   2. The server replies with a 64-byte session key and the one algorithm it picked.
///   3. The client sends a second Login with the password MAC'd against that key.
///
/// Per github.com/fogWraith/Hotline/blob/main/Docs/Protocol/HOPE-Secure-Login.md.
/// </summary>
public static class HotlineHope
{
    /// <summary>What this client calls itself. HOPE is the only point in the protocol where a client can say.</summary>
    public const string AppId = "INVG";

    public static string AppString => $"Invigoration {AppVersion.Current}";

    /// <summary>
    /// The MACs this client can compute, strongest first — the server picks the first one it also
    /// supports. INVERSE is the classic bitwise-NOT and must always be offered last as the
    /// fallback both sides are required to understand.
    /// </summary>
    public static readonly IReadOnlyList<string> SupportedMacAlgorithms =
    [
        "HMAC-SHA256",
        "HMAC-SHA1",
        "SHA1",
        "HMAC-MD5",
        "MD5",
        "INVERSE",
    ];

    /// <summary>Computes the MAC a server expects for a secret, or null for an algorithm this client doesn't know.</summary>
    public static byte[]? Mac(string algorithm, string secret, ReadOnlySpan<byte> sessionKey)
    {
        var secretBytes = Encoding.UTF8.GetBytes(secret);

        // The keyed forms MAC the session key under the secret; the bare digests hash the two
        // concatenated. Getting those the wrong way round produces a plausible-looking value that
        // every server rejects, so they're kept deliberately distinct here.
        switch (algorithm.ToUpperInvariant())
        {
            case "HMAC-SHA256":
                return HMACSHA256.HashData(secretBytes, sessionKey);
            case "HMAC-SHA1":
                return HMACSHA1.HashData(secretBytes, sessionKey);
            case "HMAC-MD5":
                return HMACMD5.HashData(secretBytes, sessionKey);
            case "SHA1":
                return SHA1.HashData(Concat(secretBytes, sessionKey));
            case "MD5":
                return MD5.HashData(Concat(secretBytes, sessionKey));
            case "INVERSE":
                return HotlineTransactionClient.XorObfuscate(secret);
            default:
                return null;
        }
    }

    private static byte[] Concat(ReadOnlySpan<byte> first, ReadOnlySpan<byte> second)
    {
        var combined = new byte[first.Length + second.Length];
        first.CopyTo(combined);
        second.CopyTo(combined.AsSpan(first.Length));
        return combined;
    }

    /// <summary>Packs a list of names as the count-then-length-prefixed form HOPE uses for algorithms, ciphers and compression.</summary>
    public static byte[] EncodeNameList(IReadOnlyList<string> names)
    {
        var encoded = names.Select(Encoding.ASCII.GetBytes).ToArray();
        var buffer = new byte[2 + encoded.Sum(n => 1 + n.Length)];
        BinaryPrimitives.WriteUInt16BigEndian(buffer, (ushort)encoded.Length);

        var offset = 2;
        foreach (var name in encoded)
        {
            buffer[offset] = (byte)name.Length;
            name.CopyTo(buffer.AsSpan(offset + 1));
            offset += 1 + name.Length;
        }

        return buffer;
    }

    /// <summary>Unpacks that same form. A server's reply carries exactly one name, but the shape is the list's.</summary>
    public static IReadOnlyList<string> DecodeNameList(ReadOnlySpan<byte> data)
    {
        // A server replying with a single algorithm may send the bare name with no count at all,
        // which is indistinguishable from a list only by whether the count makes sense. Treat
        // anything that doesn't parse as a list as one plain name.
        if (data.Length < 2)
        {
            return data.Length == 0 ? [] : [Encoding.ASCII.GetString(data)];
        }

        var count = BinaryPrimitives.ReadUInt16BigEndian(data);
        var names = new List<string>(Math.Min((int)count, 16));
        var offset = 2;
        for (var i = 0; i < count; i++)
        {
            if (offset >= data.Length)
            {
                break;
            }

            var length = data[offset];
            if (offset + 1 + length > data.Length)
            {
                break;
            }

            names.Add(Encoding.ASCII.GetString(data.Slice(offset + 1, length)));
            offset += 1 + length;
        }

        return names.Count > 0 ? names : [Encoding.ASCII.GetString(data).Trim('\0')];
    }

    /// <summary>The fields for step 1 — the null login that asks a server whether it speaks HOPE, and says what this client is.</summary>
    public static HotlineField[] IdentificationFields() =>
    [
        new(HotlineFieldType.UserLogin, new byte[] { 0x00 }),
        new(HotlineFieldType.UserPassword, new byte[] { 0x00 }),
        new(HotlineFieldType.HopeMacAlgorithm, EncodeNameList(SupportedMacAlgorithms)),
        new(HotlineFieldType.HopeAppId, Encoding.ASCII.GetBytes(AppId)),
        new(HotlineFieldType.HopeAppString, AppString),
    ];

    /// <summary>What the server said in step 2, or null when the reply isn't a HOPE one — which is how a classic server answers.</summary>
    public sealed record ServerIdentification(
        byte[] SessionKey,
        string MacAlgorithm,
        bool LoginIsMacd,
        string? ServerApp)
    {
        /// <summary>The address the server believes it is, read out of the session key. A mismatch with where we actually connected means something is in the middle — a NAT, a proxy, or an attacker.</summary>
        public (IPAddress Address, int Port)? EmbeddedEndpoint => SessionKey.Length >= 6
            ? (new IPAddress(SessionKey.AsSpan(0, 4).ToArray()), BinaryPrimitives.ReadUInt16BigEndian(SessionKey.AsSpan(4)))
            : null;
    }

    /// <summary>
    /// Reads the server's step-2 reply. Null when it carries no session key at all, meaning the
    /// server answered the null login as an ordinary (failed) one and doesn't speak HOPE.
    /// </summary>
    public static ServerIdentification? ReadServerIdentification(HotlineTransactionFrame? reply)
    {
        if (reply?.Field(HotlineFieldType.HopeSessionKey) is not { Data.Length: >= 6 } sessionKey)
        {
            return null;
        }

        var algorithms = reply.Field(HotlineFieldType.HopeMacAlgorithm) is { } macField
            ? DecodeNameList(macField.Data)
            : [];

        // An empty login field means "send the login the classic way"; a non-empty one names the
        // algorithm to MAC the login with too.
        var loginField = reply.Field(HotlineFieldType.UserLogin);
        var loginIsMacd = loginField is { Data.Length: > 0 } && loginField.Data.Any(b => b != 0);

        return new ServerIdentification(
            sessionKey.Data,
            algorithms.Count > 0 ? algorithms[0] : "INVERSE",
            loginIsMacd,
            reply.Field(HotlineFieldType.HopeAppString)?.AsString());
    }

    /// <summary>
    /// The fields for step 3 — the real login, with the password MAC'd against the server's
    /// challenge. Null when the chosen algorithm is one this client can't compute, which leaves
    /// the caller to fall back to a classic login rather than sending something meaningless.
    /// </summary>
    public static HotlineField[]? AuthenticatedLoginFields(
        ServerIdentification identification,
        string login,
        string password,
        string nickname,
        ushort iconId,
        ushort? clientVersion)
    {
        var passwordMac = Mac(identification.MacAlgorithm, password, identification.SessionKey);
        if (passwordMac is null)
        {
            return null;
        }

        // An empty password is sent as an empty field, not a MAC of nothing — matching what a real
        // client does for an anonymous login.
        var loginBytes = identification.LoginIsMacd
            ? Mac(identification.MacAlgorithm, login, identification.SessionKey) ?? HotlineTransactionClient.XorObfuscate(login)
            : HotlineTransactionClient.XorObfuscate(login);

        var fields = new List<HotlineField>
        {
            new(HotlineFieldType.UserLogin, loginBytes),
            new(HotlineFieldType.UserPassword, string.IsNullOrEmpty(password) ? [] : passwordMac),
            new(HotlineFieldType.UserIconId, iconId),
            new(HotlineFieldType.UserName, nickname),
        };

        if (clientVersion.HasValue)
        {
            fields.Add(new HotlineField(HotlineFieldType.VersionNumber, clientVersion.Value));
        }

        return [.. fields];
    }
}
