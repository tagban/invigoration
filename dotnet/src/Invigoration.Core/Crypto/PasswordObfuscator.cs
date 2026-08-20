using System.Text;

namespace Invigoration.Core.Crypto;

/// <summary>
/// Lightweight, reversible obfuscation for BotConfig.Password at rest in
/// bots.json — not real security (XOR with a fixed, published key,
/// Base64-encoded), just enough to keep a password from being immediately
/// readable by anyone who opens the file. Wrapped in brackets so it can be
/// told apart from a plaintext password typed directly into the file by
/// hand, which <see cref="Unwrap"/> passes through as-is (so a manual edit
/// works immediately) and the next save re-wraps.
/// </summary>
public static class PasswordObfuscator
{
    private static readonly byte[] Key = Encoding.UTF8.GetBytes("Invigoration");

    /// <summary>Bracket-wrapped ("[...]") text is treated as obfuscated and decoded; anything else is returned unchanged.</summary>
    public static string Unwrap(string stored)
    {
        if (stored.Length < 2 || stored[0] != '[' || stored[^1] != ']')
        {
            return stored;
        }

        try
        {
            return Encoding.UTF8.GetString(Xor(Convert.FromBase64String(stored[1..^1])));
        }
        catch (FormatException)
        {
            // Not actually obfuscated text, just a password that happens to be bracket-wrapped — use as typed.
            return stored;
        }
    }

    /// <summary>Wraps a plaintext password as "[base64]" for storage.</summary>
    public static string Wrap(string plaintext) =>
        plaintext.Length == 0 ? plaintext : $"[{Convert.ToBase64String(Xor(Encoding.UTF8.GetBytes(plaintext)))}]";

    private static byte[] Xor(byte[] bytes)
    {
        var result = new byte[bytes.Length];
        for (var i = 0; i < bytes.Length; i++)
        {
            result[i] = (byte)(bytes[i] ^ Key[i % Key.Length]);
        }

        return result;
    }
}
