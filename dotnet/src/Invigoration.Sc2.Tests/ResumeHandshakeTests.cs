using System.Text;
using Invigoration.Sc2.Native;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Tests;

/// <summary>
/// No golden hex vector exists for these records (upstream encodes them via
/// its reflective schema codec, not a hand-rolled byte sequence with its own
/// unit test). These instead verify the encoder/decoder against each other
/// and against the field-level schema pulled from the SC2Docs site and
/// protocol.rs's resume_request/proof_response/enable_encryption functions.
/// </summary>
public class ResumeHandshakeTests
{
    [Fact]
    public void EncodeResumeRequest_ProducesExpectedFieldsWhenManuallyDecoded()
    {
        var record = ResumeHandshake.EncodeResumeRequest("player@example.com", gameAccountRegion: 1, "Tagban");

        var reader = new BitReader(record);
        var routing = RoutingHeader.Decode(reader);
        Assert.Equal(1, routing.CommandId);
        Assert.Equal(ResumeHandshake.AuthenticationSlot, routing.ServiceSlot);

        Assert.Equal(FourCc.Encode("S2"), (uint)reader.Read(32));
        Assert.Equal(FourCc.Encode("Mc64"), (uint)reader.Read(32));
        Assert.Equal(FourCc.Encode("enUS"), (uint)reader.Read(32));

        Assert.Equal(5u, reader.Read(6));
        for (var i = 0; i < 4; i++)
        {
            reader.Read(32);
            reader.Read(32);
            reader.Read(32);
        }

        Assert.Equal(FourCc.Encode("Bnet"), (uint)reader.Read(32));
        Assert.Equal(FourCc.Encode("Mc64"), (uint)reader.Read(32));
        Assert.Equal(0x000a16a7u, (uint)reader.Read(32));

        var accountLength = (int)reader.Read(9);
        var accountBytes = reader.ReadBytes(accountLength, aligned: true);
        Assert.Equal("player@example.com", Encoding.UTF8.GetString(accountBytes));

        Assert.Equal(1u, reader.Read(8));

        var nameLength = (int)reader.Read(5);
        var nameBytes = reader.ReadBytes(nameLength, aligned: true);
        Assert.Equal("Tagban", Encoding.UTF8.GetString(nameBytes));
    }

    [Fact]
    public void EncodeEnableEncryption_IsTwoBytesWithNoPayload()
    {
        var record = ResumeHandshake.EncodeEnableEncryption();

        Assert.Equal(2, record.Length);
        var reader = new BitReader(record);
        var routing = RoutingHeader.Decode(reader);
        Assert.Equal(5, routing.CommandId);
        Assert.Equal(ResumeHandshake.ConnectionSlot, routing.ServiceSlot);
    }

    [Fact]
    public void EncodeProofResponse_WrapsSessionProofAsSingleModuleOutput()
    {
        var sessionSeed = new byte[64];
        for (var i = 0; i < sessionSeed.Length; i++)
        {
            sessionSeed[i] = (byte)i;
        }

        var serverNonce = new byte[16];
        var clientNonce = new byte[16];
        for (var i = 0; i < 16; i++)
        {
            clientNonce[i] = (byte)(i + 1);
        }

        var proof = NativeCrypto.BuildSessionProofWithNonce(sessionSeed, serverNonce, clientNonce);

        var record = ResumeHandshake.EncodeProofResponse(proof.Output);

        var reader = new BitReader(record);
        var routing = RoutingHeader.Decode(reader);
        Assert.Equal(2, routing.CommandId);
        Assert.Equal(ResumeHandshake.AuthenticationSlot, routing.ServiceSlot);
        Assert.Equal(1u, reader.Read(3));
        var length = (int)reader.Read(10);
        Assert.Equal(49, length);
        var data = reader.ReadBytes(length, aligned: true);
        Assert.Equal(proof.Output, data);
    }

    [Fact]
    public void DecodeConfiguration_ReadsUseS3DepotFlag()
    {
        var writer = new BitWriter();
        writer.Write(1, 1);

        var config = ResumeHandshake.DecodeConfiguration(new BitReader(writer.ToBytes()));

        Assert.True(config.UseS3Depot);
    }

    [Fact]
    public void DecodeProofRequest_ParsesTwoModules()
    {
        var writer = new BitWriter();
        writer.Write(2, 3);
        for (var i = 0; i < 2; i++)
        {
            writer.WriteBytes("auth\0\0\0\0"u8, aligned: true);
            writer.WriteBytes(new byte[32], aligned: true);
            writer.Write(0, 10);
        }

        var modules = ResumeHandshake.DecodeProofRequest(new BitReader(writer.ToBytes()));

        Assert.Equal(2, modules.Count);
        Assert.Equal(8, modules[0].Id.Usage.Length);
        Assert.Equal(32, modules[0].Id.Identity.Length);
        Assert.Empty(modules[0].Data);
    }

    [Fact]
    public void DecodeResumeResponse_Success_ParsesPingTimeoutAndNoRegulator()
    {
        var writer = new BitWriter();
        writer.Write(0, 1); // success
        writer.Write(0, 3); // no final requests
        writer.Write((uint)30, 32); // ping timeout
        writer.Write(0, 1); // regulator rules absent

        var result = ResumeHandshake.DecodeResumeResponse(new BitReader(writer.ToBytes()));

        var success = Assert.IsType<ResumeResult.Success>(result);
        Assert.Empty(success.FinalRequests);
        Assert.Equal(30, success.PingTimeoutSeconds);
        Assert.Null(success.RegulatorRules);
    }

    [Fact]
    public void DecodeResumeResponse_Success_ParsesLeakyBucketRegulator()
    {
        var writer = new BitWriter();
        writer.Write(0, 1); // success
        writer.Write(0, 3); // no final requests
        writer.Write((uint)15, 32); // ping timeout
        writer.Write(1, 1); // regulator rules present
        writer.Write(1, 1); // Info selector: LeakyBucket
        writer.Write(1000, 32); // threshold
        writer.Write(10, 32); // rate

        var result = ResumeHandshake.DecodeResumeResponse(new BitReader(writer.ToBytes()));

        var success = Assert.IsType<ResumeResult.Success>(result);
        var bucket = Assert.IsType<RegulatorInfo.LeakyBucket>(success.RegulatorRules);
        Assert.Equal(1000u, bucket.Threshold);
        Assert.Equal(10u, bucket.Rate);
    }

    [Fact]
    public void DecodeResumeResponse_Failure_ParsesErrorCodeAndWait()
    {
        var writer = new BitWriter();
        writer.Write(1, 1); // failure
        writer.Write(0, 1); // strings absent
        writer.Write(1, 2); // reason selector: failure
        writer.Write(1234, 16); // error code
        writer.Write((uint)5, 32); // wait seconds

        var result = ResumeHandshake.DecodeResumeResponse(new BitReader(writer.ToBytes()));

        var failure = Assert.IsType<ResumeResult.Failure>(result);
        Assert.Null(failure.Strings);
        var reason = Assert.IsType<ResumeFailureReason.Failed>(failure.Reason);
        Assert.Equal(1234, reason.ErrorCode);
        Assert.Equal(5, reason.WaitSeconds);
    }

    [Fact]
    public void DecodeResumeResponse_Failure_VersionCheckDisconnect()
    {
        var writer = new BitWriter();
        writer.Write(1, 1); // failure
        writer.Write(0, 1); // strings absent
        writer.Write(2, 2); // reason selector: versionCheckDisconnect

        var result = ResumeHandshake.DecodeResumeResponse(new BitReader(writer.ToBytes()));

        var failure = Assert.IsType<ResumeResult.Failure>(result);
        Assert.IsType<ResumeFailureReason.VersionCheckDisconnect>(failure.Reason);
    }
}

/// <summary>
/// AnswerProofRequest's happy path requires a real RSA signature over the
/// thumbprint challenge, which only Blizzard's private key can produce —
/// the reference crate's own test suite has the same gap. These cover every
/// error path instead, which don't need one.
/// </summary>
public class AnswerProofRequestTests
{
    private static readonly byte[] Usage = "auth\0\0\0\0"u8.ToArray();

    private static readonly byte[] ThumbprintIdentity =
    [
        0xd7, 0xe6, 0x62, 0x40, 0x80, 0xc1, 0xab, 0xa6, 0x6d, 0xee, 0x63, 0xa6, 0xf3, 0x92, 0x8d, 0x8a,
        0x54, 0x69, 0x25, 0x7f, 0x58, 0x20, 0xb5, 0x72, 0x1f, 0xb8, 0xc3, 0x2b, 0x6b, 0x5b, 0xef, 0x5d,
    ];

    private static readonly byte[] SessionProofIdentity =
    [
        0x89, 0x50, 0x05, 0x34, 0x0a, 0x63, 0x0a, 0x64, 0x65, 0xa6, 0x5f, 0xec, 0x96, 0x32, 0x3c, 0x31,
        0x0b, 0xca, 0x8a, 0x9f, 0x66, 0xec, 0xee, 0xb1, 0x88, 0x7a, 0x9d, 0x6c, 0x0e, 0x67, 0x61, 0x2e,
    ];

    private static readonly byte[] SessionSeed = Enumerable.Range(0, 64).Select(i => (byte)i).ToArray();

    private static ModuleInput ValidSessionModule() =>
        new(new ModuleId(Usage, SessionProofIdentity), [0, .. new byte[16]]);

    [Fact]
    public void EmptyModuleList_MissingRequiredModule_Throws()
    {
        var ex = Assert.Throws<InvalidOperationException>(() =>
            ResumeHandshake.AnswerProofRequest([], SessionSeed, new byte[16]));
        Assert.Contains("missing a required auth module", ex.Message);
    }

    [Fact]
    public void UnknownModule_Throws()
    {
        var unknown = new ModuleInput(new ModuleId(Usage, new byte[32]), []);

        Assert.Throws<InvalidOperationException>(() =>
            ResumeHandshake.AnswerProofRequest([unknown], SessionSeed, new byte[16]));
    }

    [Fact]
    public void RepeatedSessionModule_Throws()
    {
        var modules = new[] { ValidSessionModule(), ValidSessionModule() };

        var ex = Assert.Throws<InvalidOperationException>(() =>
            ResumeHandshake.AnswerProofRequest(modules, SessionSeed, new byte[16]));
        Assert.Contains("repeats the session module", ex.Message);
    }

    [Fact]
    public void InvalidSessionPhase_Throws()
    {
        var badPhase = new ModuleInput(new ModuleId(Usage, SessionProofIdentity), [1, .. new byte[16]]);

        var ex = Assert.Throws<InvalidOperationException>(() =>
            ResumeHandshake.AnswerProofRequest([badPhase], SessionSeed, new byte[16]));
        Assert.Contains("session-proof phase", ex.Message);
    }

    [Fact]
    public void ThumbprintModuleWithFakeSignature_Rejected()
    {
        var thumbprint = new ModuleInput(new ModuleId(Usage, ThumbprintIdentity), new byte[512]);

        var ex = Assert.Throws<InvalidOperationException>(() =>
            ResumeHandshake.AnswerProofRequest([thumbprint, ValidSessionModule()], SessionSeed, new byte[16]));
        Assert.Contains("thumbprint proof failed", ex.Message);
    }
}

public class ValidateServerProofTests
{
    private static readonly byte[] Usage = "auth\0\0\0\0"u8.ToArray();

    private static readonly byte[] SessionProofIdentity =
    [
        0x89, 0x50, 0x05, 0x34, 0x0a, 0x63, 0x0a, 0x64, 0x65, 0xa6, 0x5f, 0xec, 0x96, 0x32, 0x3c, 0x31,
        0x0b, 0xca, 0x8a, 0x9f, 0x66, 0xec, 0xee, 0xb1, 0x88, 0x7a, 0x9d, 0x6c, 0x0e, 0x67, 0x61, 0x2e,
    ];

    private static NativeCrypto.SessionProof BuildProof() => NativeCrypto.BuildSessionProofWithNonce(
        Enumerable.Range(0, 64).Select(i => (byte)i).ToArray(),
        Enumerable.Range(16, 16).Select(i => (byte)i).ToArray(),
        Enumerable.Range(32, 16).Select(i => (byte)i).ToArray());

    [Fact]
    public void CorrectProof_Succeeds()
    {
        var proof = BuildProof();
        var success = new ResumeResult.Success(
            [new ModuleInput(new ModuleId(Usage, SessionProofIdentity), [2, .. proof.ExpectedServerProof])],
            PingTimeoutSeconds: 30,
            RegulatorRules: null);

        ResumeHandshake.ValidateServerProof(success, proof);
    }

    [Fact]
    public void WrongProofBytes_Throws()
    {
        var proof = BuildProof();
        var wrongProof = new byte[32];
        var success = new ResumeResult.Success(
            [new ModuleInput(new ModuleId(Usage, SessionProofIdentity), [2, .. wrongProof])],
            PingTimeoutSeconds: 30,
            RegulatorRules: null);

        Assert.Throws<InvalidOperationException>(() => ResumeHandshake.ValidateServerProof(success, proof));
    }

    [Fact]
    public void WrongFinalRequestCount_Throws()
    {
        var proof = BuildProof();
        var success = new ResumeResult.Success([], PingTimeoutSeconds: 30, RegulatorRules: null);

        Assert.Throws<InvalidOperationException>(() => ResumeHandshake.ValidateServerProof(success, proof));
    }

    [Fact]
    public void WrongPhaseByte_Throws()
    {
        var proof = BuildProof();
        var success = new ResumeResult.Success(
            [new ModuleInput(new ModuleId(Usage, SessionProofIdentity), [3, .. proof.ExpectedServerProof])],
            PingTimeoutSeconds: 30,
            RegulatorRules: null);

        Assert.Throws<InvalidOperationException>(() => ResumeHandshake.ValidateServerProof(success, proof));
    }
}
