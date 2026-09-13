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
/// on a private server and broke a bot-made one on live Battle.net (2026-09-13). The logon now
/// sends the password as typed first — unchanged for every account that already worked — and, if
/// the server rejects it as incorrect, retries once with <see cref="RetryCasing"/>.</para>
/// </remarks>
public static class BattlenetPassword
{
    /// <summary>The form Blizzard's game clients hash: uppercase for WC3/TFT, lowercase for everything else.</summary>
    public static string Normalize(string password, string product) =>
        product is BncsProduct.Warcraft3 or BncsProduct.Warcraft3TFT ? password.ToUpperInvariant() : password.ToLowerInvariant();

    /// <summary>The form to retry with after the password as typed was rejected, or null when it's already in that form (a retry would just fail the same way).</summary>
    public static string? RetryCasing(string typed, string product) =>
        Normalize(typed, product) is var normalized && normalized != typed ? normalized : null;
}
