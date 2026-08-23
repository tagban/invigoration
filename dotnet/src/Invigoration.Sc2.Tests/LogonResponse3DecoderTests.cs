using Invigoration.Sc2.Native;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Tests;

/// <summary>
/// No captured packet exists for LogonResponse3 either — this is a
/// synthetic round-trip against the SC2Docs schema, same caveat as the
/// other unvectored decoders in this project.
/// </summary>
public class LogonResponse3DecoderTests
{
    [Fact]
    public void Decode_Success_ParsesGameAccountRegionSeparatelyFromAccountRegion()
    {
        var writer = new BitWriter();
        writer.Write(0, 1); // Logon: 0 bits (nothing written) + m_result selector: success
        writer.Write(0, 3); // ResponseSuccessCommon.m_finalRequest: 0 modules
        // Battlenet::s32 wire value is raw + minimum (minimum = -2^31), i.e. the sign bit flipped
        // relative to a plain two's-complement bit-cast. See LogonResponse3Decoder's pingTimeoutSeconds comment.
        writer.Write(30u ^ 0x8000_0000u, 32); // m_pingTimeout
        writer.Write(0, 1); // m_regulatorRules: absent
        writer.Write(0, 6); // m_givenName: 0 bytes
        writer.WriteBytes([], aligned: true); // NamePart is byte-aligned even when empty
        writer.Write(0, 6); // m_surname: 0 bytes
        writer.WriteBytes([], aligned: true);
        writer.Write(1149051, 32); // m_accountId
        writer.Write(0, 8); // m_accountRegion
        writer.Write(0UL, 64); // m_accountFlags
        writer.Write(3, 8); // m_gameAccountRegion -- deliberately different from m_accountRegion
        var name = "1149051#1"u8.ToArray();
        writer.Write((ulong)(name.Length - 1), 5); // biased -1
        writer.WriteBytes(name, aligned: true);
        writer.Write(0UL, 64); // m_gameAccountFlags
        writer.Write(0, 32); // m_logonFailures
        writer.Align();

        var result = LogonResponse3Decoder.Decode(writer.ToBytes());

        Assert.Equal(1149051u, result.AccountId);
        Assert.Equal(0, result.AccountRegion);
        Assert.Equal(3, result.GameAccountRegion);
        Assert.Equal("1149051#1", result.GameAccountName);
        Assert.Equal(30, result.PingTimeoutSeconds);
    }

    [Fact]
    public void Decode_Failure_Throws()
    {
        var writer = new BitWriter();
        writer.Write(1, 1); // m_result selector: failure

        Assert.Throws<InvalidOperationException>(() => LogonResponse3Decoder.Decode(writer.ToBytes()));
    }
}
