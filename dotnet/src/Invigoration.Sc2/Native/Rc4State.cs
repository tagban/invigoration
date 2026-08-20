namespace Invigoration.Sc2.Native;

/// <summary>
/// Standard RC4 (KSA + PRGA, no key-drop), with state persisting across
/// calls — SC2's native transport keeps one long-lived keystream per
/// direction for the life of the connection rather than re-keying per
/// message. Ported from core/src/native/crypto.rs's Rc4State.
/// </summary>
public sealed class Rc4State
{
    private readonly byte[] _state = new byte[256];
    private byte _i;
    private byte _j;

    public Rc4State(ReadOnlySpan<byte> key)
    {
        if (key.Length is < 1 or > 256)
        {
            throw new ArgumentOutOfRangeException(nameof(key), "RC4 key length must be 1..=256 bytes.");
        }

        for (var i = 0; i < 256; i++)
        {
            _state[i] = (byte)i;
        }

        byte j = 0;
        for (var i = 0; i < 256; i++)
        {
            j = (byte)(j + _state[i] + key[i % key.Length]);
            (_state[i], _state[j]) = (_state[j], _state[i]);
        }
    }

    public void ApplyInPlace(Span<byte> data)
    {
        for (var n = 0; n < data.Length; n++)
        {
            _i++;
            _j = (byte)(_j + _state[_i]);
            (_state[_i], _state[_j]) = (_state[_j], _state[_i]);
            data[n] ^= _state[(byte)(_state[_i] + _state[_j])];
        }
    }

    public byte[] Apply(ReadOnlySpan<byte> data)
    {
        var output = data.ToArray();
        ApplyInPlace(output);
        return output;
    }
}
