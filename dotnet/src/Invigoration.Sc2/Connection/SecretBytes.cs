namespace Invigoration.Sc2.Connection;

/// <summary>
/// A validated Battle.net web-auth credential ("ST" token) extracted from
/// the login redirect URL. Mirrors core/src/bgs (SecretBytes) — the same
/// bounds the reference client's redirect parser enforces.
/// </summary>
public sealed class SecretBytes
{
    public byte[] Value { get; }

    private SecretBytes(byte[] value) => Value = value;

    public static SecretBytes? TryCreate(byte[] value)
    {
        if (value.Length == 0 || value.Length > 1024 || !value.All(IsAsciiGraphic))
        {
            return null;
        }

        return new SecretBytes(value);
    }

    private static bool IsAsciiGraphic(byte b) => b is >= 0x21 and <= 0x7e;
}
