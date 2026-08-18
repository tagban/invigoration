using System.Text;

namespace Invigoration.Core.Auth;

/// <summary>
/// Computes the CRC-32 challenge response BNLS_AUTHORIZE expects: CRC32 of the
/// shared secret followed by the 8-digit uppercase hex server code. Port of
/// modBNLS.bas's BNLSChecksum.
/// </summary>
public static class BnlsChecksum
{
    public static uint Compute(string sharedSecret, uint serverCode)
    {
        var text = sharedSecret + serverCode.ToString("X8");
        return Crc32.Compute(Encoding.Latin1.GetBytes(text));
    }
}
