using System.Buffers.Binary;
using System.Text;
using Invigoration.Core.Protocol;

namespace Invigoration.Core.Auth;

/// <summary>
/// Which letter case a Battle.net password is sent in. A server only ever checks against the hash
/// it stored when the account was made, so the right case depends on who made the account, not
/// on the server:
/// </summary>
/// <remarks>
/// <para>Blizzard's own game clients normalize before hashing — lowercase, or uppercase for
/// Warcraft III and The Frozen Throne, whose NLS/SRP logon hashes it uppercased. An account
/// created from a real client (or a private server's web signup, which does the same) only
/// accepts that form.</para>
/// <para>This bot has always sent passwords exactly as typed, including when it creates the
/// account itself, so any account it made — on official Battle.net or anywhere else — has the
/// as-typed hash on file and only accepts that form.</para>
/// <para>So neither form works everywhere. Forcing the normalized form fixed a client-made account
/// on a private server and broke a bot-made one on live Battle.net (2026-09-13).</para>
/// <para>The logon starts in the game clients' casing (<see cref="FirstAttempt"/>), which every
/// client-made account and every account this bot creates now accepts. If Battle.net calls it
/// incorrect, it tries once more in the other form (<see cref="OtherCasing"/>) on a fresh
/// connection — a whole new handshake, so there's no doubt the server really checked it. Whichever
/// form gets in is remembered on the bot (BotConfig.PasswordSentAsTyped), so an older bot-made
/// account pays for the extra attempt once, not on every logon — failed logons add up to lockouts.</para>
/// </remarks>
public static class BattlenetPassword
{
    /// <summary>The form Blizzard's game clients hash: uppercase for WC3/TFT, lowercase for everything else.</summary>
    public static string Normalize(string password, string product) =>
        product is BncsProduct.Warcraft3 or BncsProduct.Warcraft3TFT ? password.ToUpperInvariant() : password.ToLowerInvariant();

    /// <summary>The game clients' casing of a password, or null when it's already in that form — nothing for Normalize Password to change.</summary>
    public static string? RetryCasing(string typed, string product) =>
        Normalize(typed, product) is var normalized && normalized != typed ? normalized : null;

    /// <summary>What a logon sends first: the game clients' casing, or the password exactly as typed for an account known to need that.</summary>
    public static string FirstAttempt(string typed, string product, bool sentAsTyped) =>
        sentAsTyped ? typed : Normalize(typed, product);

    /// <summary>What to try after <see cref="FirstAttempt"/> was rejected, or null when both forms are the same string (it could only fail the same way).</summary>
    public static string? OtherCasing(string typed, string product, bool sentAsTyped) =>
        RetryCasing(typed, product) is { } normalized ? (sentAsTyped ? normalized : typed) : null;

    /// <summary>
    /// The classic logon's password hash: X-SHA1 of the password's bytes, exactly as given (no
    /// case change here — see <see cref="RetryCasing"/>). Computed locally, the same way the CD key
    /// is hashed for SID_AUTH_CHECK, instead of asked of BNLS: a BNLS_HASHDATA request carries the
    /// password itself, in plain text, to a third-party server. What SID_CREATEACCOUNT and the new
    /// password in SID_CHANGEPASSWORD carry.
    /// </summary>
    public static byte[] Hash(string password) => XSha1.Hash(Encoding.Latin1.GetBytes(password));

    /// <summary>What proves the password for one handshake: X-SHA1 of the client token, the server token and <see cref="Hash"/>. Sent by SID_LOGONRESPONSE2, SID_LOGONREALMEX and (for the old password) SID_CHANGEPASSWORD.</summary>
    public static byte[] Proof(uint clientToken, uint serverToken, string password)
    {
        var tokens = new byte[8];
        BinaryPrimitives.WriteUInt32LittleEndian(tokens, clientToken);
        BinaryPrimitives.WriteUInt32LittleEndian(tokens.AsSpan(4), serverToken);
        return XSha1.Hash(tokens, Hash(password));
    }
}
