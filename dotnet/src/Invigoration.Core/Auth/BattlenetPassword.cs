using Invigoration.Core.Protocol;

namespace Invigoration.Core.Auth;

/// <summary>
/// Classic Battle.net account passwords are case-insensitive: the game clients lowercase a
/// password before hashing it, so the hash an account has on file is always of the lowercase
/// form. A password sent with its original capitals hashes to something the server never stored,
/// and the logon fails as "incorrect password" even though the user typed it exactly right.
/// </summary>
/// <remarks>
/// <para>Warcraft III and The Frozen Throne are the exception: they log on through the NLS/SRP
/// system, which hashes the password uppercased instead. Uppercasing here is also safe if a BNLS
/// server already does it itself, since doing it twice changes nothing.</para>
/// <para>Applied at the point a password leaves the bot (every BNLS hash, logon-challenge and
/// account-create request) rather than to BotConfig.Password itself, so the Config window still
/// shows what the user actually typed.</para>
/// </remarks>
public static class BattlenetPassword
{
    public static string Normalize(string password, string product) =>
        product is BncsProduct.Warcraft3 or BncsProduct.Warcraft3TFT ? password.ToUpperInvariant() : password.ToLowerInvariant();
}
