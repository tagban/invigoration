using System.Text;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

/// <summary>
/// The fields of a successful Battlenet::Client::Authentication::LogonResponse3
/// actually needed to drive Resume. This is a bit-packed native structure
/// smuggled inside the Front GameUtilities handoff's opaque "logon_response"
/// blob attribute (<see cref="Front.SunkenHandoff.LogonResponse"/>) — despite
/// arriving over the Front/protobuf transport, its own encoding is the same
/// BSN bit-packed format as everything else in this namespace.
/// </summary>
public sealed record LogonResponse3Success(
    uint AccountId,
    byte AccountRegion,
    ulong AccountFlags,
    byte GameAccountRegion,
    string GameAccountName,
    ulong GameAccountFlags,
    uint LogonFailures,
    int PingTimeoutSeconds);

/// <summary>
/// Decodes LogonResponse3. Critically, this is where the real
/// <c>game_account_region</c> for <see cref="ResumeHandshake.EncodeResumeRequest"/>
/// comes from — NOT the top-level GameUtilities "account_region" attribute,
/// which is a different, only-used-for-a-consistency-check field. Confirmed
/// directly from core/src/native/protocol.rs's decode_logon_parameters,
/// which is what actually supplies the value Resume sends upstream.
/// </summary>
public static class LogonResponse3Decoder
{
    public static LogonResponse3Success Decode(byte[] logonResponseBlob)
    {
        var reader = new BitReader(logonResponseBlob);

        // Logon: 0 fields / 0 bits.
        var isFailure = reader.Read(1) != 0;
        if (isFailure)
        {
            throw new InvalidOperationException("Front native logon was rejected.");
        }

        // ResponseSuccessCommon
        var finalRequestCount = (int)reader.Read(3);
        for (var i = 0; i < finalRequestCount; i++)
        {
            SkipModuleInput(reader);
        }

        var pingTimeoutSeconds = unchecked((int)reader.Read(32));
        if (reader.Read(1) != 0 && reader.Read(1) != 0) // regulatorRules present, and it's LeakyBucket
        {
            reader.Read(32);
            reader.Read(32);
        }

        // m_fullName
        SkipNamePart(reader); // m_givenName
        SkipNamePart(reader); // m_surname

        var accountId = (uint)reader.Read(32);
        var accountRegion = (byte)reader.Read(8);
        var accountFlags = reader.Read(64);
        var gameAccountRegion = (byte)reader.Read(8);

        var gameAccountNameByteCount = (int)reader.Read(5) + 1; // biased +1: raw 0..31 -> 1..32 bytes
        var gameAccountNameBytes = reader.ReadBytes(gameAccountNameByteCount, aligned: true);
        var gameAccountName = Encoding.UTF8.GetString(gameAccountNameBytes);

        var gameAccountFlags = reader.Read(64);
        var logonFailures = (uint)reader.Read(32);

        return new LogonResponse3Success(
            accountId,
            accountRegion,
            accountFlags,
            gameAccountRegion,
            gameAccountName,
            gameAccountFlags,
            logonFailures,
            pingTimeoutSeconds);
    }

    private static void SkipModuleInput(BitReader reader)
    {
        reader.ReadBytes(40, aligned: true); // ModuleId: 8-byte usage + 32-byte identity
        var dataLength = (int)reader.Read(10);
        reader.ReadBytes(dataLength, aligned: true);
    }

    private static void SkipNamePart(BitReader reader)
    {
        var byteCount = (int)reader.Read(6); // no bias: accepted range 0..=32 matches raw 6-bit range's low end
        reader.ReadBytes(byteCount, aligned: true);
    }
}
