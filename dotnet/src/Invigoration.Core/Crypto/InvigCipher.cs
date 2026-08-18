using System.Text;

namespace Invigoration.Core.Crypto;

/// <summary>
/// Invigoration's homegrown chat obfuscation (not real cryptography — a
/// simple digit-shift scheme): each character becomes a 3-digit decimal
/// ASCII code, and each resulting digit character is shifted by +101. Port
/// of InvigEncrypt/InvigDecrypt in modFunctions.bas, used by the
/// "invigencrypt"/"hex" chat commands.
/// </summary>
public static class InvigCipher
{
    /// <summary>
    /// Matches the original exactly, including its quirk: an odd-length
    /// input silently drops its last character (the "invigencrypt" command
    /// works around this by appending a trailing "-" before calling in).
    /// </summary>
    public static string Encrypt(string text)
    {
        var usableLength = text.Length - (text.Length % 2);
        var digits = new StringBuilder(usableLength * 3);
        for (var i = 0; i < usableLength; i++)
        {
            digits.Append(((int)text[i]).ToString("000"));
        }

        var result = new StringBuilder(digits.Length);
        foreach (var d in digits.ToString())
        {
            result.Append((char)(d + 101));
        }

        return result.ToString();
    }

    public static string Decrypt(string text)
    {
        try
        {
            var digits = new StringBuilder(text.Length);
            foreach (var c in text)
            {
                digits.Append((char)(c - 101));
            }

            var s = digits.ToString();
            var result = new StringBuilder();
            for (var i = 0; i < s.Length; i += 3)
            {
                var chunk = s.Substring(i, Math.Min(3, s.Length - i));
                result.Append((char)int.Parse(chunk));
            }

            return result.ToString();
        }
        catch
        {
            return "<< Invigoration Decryption Failed >>";
        }
    }
}
