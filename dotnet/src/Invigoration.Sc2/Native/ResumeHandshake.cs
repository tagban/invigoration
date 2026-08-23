using System.Security.Cryptography;
using System.Text;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

/// <summary>
/// Encodes/decodes the plaintext Sunken records exchanged during Resume:
/// Auth/1 ResumeRequest, Auth/18 Configuration, Auth/2 ProofRequest/
/// ProofResponse, Auth/1 ResumeResponse, and Conn/5 EnableEncryption.
///
/// Field widths and ordering are cross-checked from two independent
/// primary sources: the field-level schema published at
/// https://superioritybot.com/PROTOCOL's type/service registries, and the
/// hand-rolled reflective-codec call sites in core/src/native/protocol.rs
/// (resume_request, proof_response, enable_encryption) — which additionally
/// confirmed the slot/command constants (AUTHENTICATION_SLOT=0,
/// CONNECTION_SLOT=1) and the exact RequestCommon version-table values.
/// Unlike the chat commands, upstream itself encodes these via its
/// reflective schema codec rather than a hand-rolled bit writer, so there is
/// no golden hex vector to test against here — these are schema-derived,
/// cross-referenced between two sources, not independently vector-verified.
///
/// Two fields' length/value prefixes are *biased*, not raw: BSN's generic
/// integer codec (core/src/bsn/codec.rs's read_range/write_range) always
/// stores `wire = value - range.minimum` in `range.bit_width` bits, not
/// `value` itself. For most fields here `minimum` is 0 so this is invisible,
/// but core/src/native/schema/wire.rs's static type table gives
/// Battlenet::Account::Mail (m_account) `IntegerRange { bit_width: 9,
/// minimum: 3, maximum: 320 }` and Battlenet::GameAccount::Name
/// (m_gameAccountName) `IntegerRange { bit_width: 5, minimum: 1, maximum: 32
/// }` — both non-zero. Writing the raw byte count instead of `count -
/// minimum` (as this file did until the Boom(6)-rejection bug was root-
/// caused against that reference) desyncs every bit after it in the record,
/// which upstream's server answers by tearing down the connection with a
/// Connection/Boom record. <see cref="LogonResponse3Decoder"/> already knew
/// about the GameAccount::Name bias on the decode side (see its "biased +1"
/// comment) — this file's own encoder just never matched it.
/// </summary>
public static class ResumeHandshake
{
    public const byte AuthenticationSlot = 0;
    public const byte ConnectionSlot = 1;

    private const byte AuthResumeCommand = 1;
    private const byte AuthProofCommand = 2;
    private const byte ConnectionEnableEncryptionCommand = 5;

    private static readonly (string Program, string Component, uint Version)[] MacosNativeVersions =
    [
        ("S2", "NGD1", 0x5bc8dcc1),
        ("S2", "NGD2", 0xfade3a32),
        ("S2", "NGD3", 0x0c129365),
        ("S2", "NGD4", 0x86b7c0ed),
        ("Bnet", "Mc64", 0x000a16a7),
    ];

    private static readonly byte[] ModuleUsage = "auth\0\0\0\0"u8.ToArray();

    private static readonly byte[] ThumbprintModuleIdentity =
    [
        0xd7, 0xe6, 0x62, 0x40, 0x80, 0xc1, 0xab, 0xa6, 0x6d, 0xee, 0x63, 0xa6, 0xf3, 0x92, 0x8d, 0x8a,
        0x54, 0x69, 0x25, 0x7f, 0x58, 0x20, 0xb5, 0x72, 0x1f, 0xb8, 0xc3, 0x2b, 0x6b, 0x5b, 0xef, 0x5d,
    ];

    private static readonly byte[] SessionProofModuleIdentity =
    [
        0x89, 0x50, 0x05, 0x34, 0x0a, 0x63, 0x0a, 0x64, 0x65, 0xa6, 0x5f, 0xec, 0x96, 0x32, 0x3c, 0x31,
        0x0b, 0xca, 0x8a, 0x9f, 0x66, 0xec, 0xee, 0xb1, 0x88, 0x7a, 0x9d, 0x6c, 0x0e, 0x67, 0x61, 0x2e,
    ];

    /// <summary>Result of answering an Auth/2 ProofRequest: the 49-byte session-proof output to send back, and the session state needed to later validate the server's own proof.</summary>
    public sealed record ProofAnswer(byte[] SessionProofOutput, NativeCrypto.SessionProof Session);

    public static byte[] EncodeResumeRequest(string accountMail, byte gameAccountRegion, string gameAccountName)
    {
        var accountBytes = Encoding.UTF8.GetBytes(accountMail);
        if (accountBytes.Length is < 3 or > 320)
        {
            throw new ArgumentException("Account mail must be 3..=320 UTF-8 bytes.", nameof(accountMail));
        }

        var gameAccountNameBytes = Encoding.UTF8.GetBytes(gameAccountName);
        if (gameAccountNameBytes.Length is < 1 or > 32)
        {
            throw new ArgumentException("Game account name must be 1..=32 UTF-8 bytes.", nameof(gameAccountName));
        }

        var writer = new BitWriter();
        RoutingHeader.Encode(writer, AuthResumeCommand, AuthenticationSlot);
        WriteRequestCommon(writer);
        // Account::Mail: IntegerRange { bit_width: 9, minimum: 3, maximum: 320 } -- biased -3.
        WriteBlob(writer, accountBytes, lengthBits: 9, minimum: 3);
        writer.Write(gameAccountRegion, 8);
        // GameAccount::Name: IntegerRange { bit_width: 5, minimum: 1, maximum: 32 } -- biased -1.
        WriteBlob(writer, gameAccountNameBytes, lengthBits: 5, minimum: 1);
        writer.Align();
        return writer.ToBytes();
    }

    /// <summary>Builds the Auth/2 ProofResponse carrying exactly one output — the session-proof module's 49-byte proof. The thumbprint module is verified locally and contributes zero outputs, per upstream.</summary>
    public static byte[] EncodeProofResponse(byte[] sessionProofOutput)
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, AuthProofCommand, AuthenticationSlot);
        writer.Write(1, 3); // responses count: one session-proof output
        WriteBlob(writer, sessionProofOutput, lengthBits: 10);
        writer.Align();
        return writer.ToBytes();
    }

    public static byte[] EncodeEnableEncryption()
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, ConnectionEnableEncryptionCommand, ConnectionSlot);
        writer.Align();
        return writer.ToBytes();
    }

    public static ResumeConfiguration DecodeConfiguration(BitReader reader)
    {
        var useS3Depot = reader.Read(1) != 0;
        return new ResumeConfiguration(useS3Depot);
    }

    public static IReadOnlyList<ModuleInput> DecodeProofRequest(BitReader reader)
    {
        var count = (int)reader.Read(3);
        var modules = new List<ModuleInput>(count);
        for (var i = 0; i < count; i++)
        {
            modules.Add(DecodeModuleInput(reader));
        }

        return modules;
    }

    /// <summary>
    /// Verifies the thumbprint module and builds the session-proof output
    /// for the (unordered) two modules in an Auth/2 ProofRequest. Mirrors
    /// core/src/native/auth.rs's answer_proof_request exactly, including
    /// its validation order and error conditions.
    /// </summary>
    public static ProofAnswer AnswerProofRequest(IReadOnlyList<ModuleInput> modules, byte[] sessionSeed, byte[] thumbprintContext)
    {
        var sawThumbprint = false;
        NativeCrypto.SessionProof? session = null;

        foreach (var module in modules)
        {
            if (IsModule(module.Id, ThumbprintModuleIdentity))
            {
                if (sawThumbprint)
                {
                    throw new InvalidOperationException("Proof request repeats the thumbprint module.");
                }

                sawThumbprint = true;
                if (!NativeCrypto.VerifyThumbprint(thumbprintContext, module.Data))
                {
                    throw new InvalidOperationException("Native server thumbprint proof failed.");
                }
            }
            else if (IsModule(module.Id, SessionProofModuleIdentity))
            {
                if (session is not null)
                {
                    throw new InvalidOperationException("Proof request repeats the session module.");
                }

                if (module.Data.Length != 17 || module.Data[0] != 0)
                {
                    throw new InvalidOperationException("Unsupported native session-proof phase.");
                }

                session = NativeCrypto.BuildSessionProof(sessionSeed, module.Data[1..]);
            }
            else
            {
                throw new InvalidOperationException("Server requested an unknown auth module.");
            }
        }

        if (!sawThumbprint || session is null)
        {
            throw new InvalidOperationException("Proof request is missing a required auth module.");
        }

        return new ProofAnswer(session.Output, session);
    }

    /// <summary>
    /// Validates the server's phase-two proof, which upstream smuggles
    /// inside a successful ResumeResponse's <c>final_requests</c> array
    /// (reusing the same {id, data} shape as an auth module request) rather
    /// than a dedicated field — confirmed directly from
    /// core/src/native/auth.rs's validate_resume_server_proof, since the
    /// published field-level schema alone doesn't make this obvious.
    /// </summary>
    public static void ValidateServerProof(ResumeResult.Success success, NativeCrypto.SessionProof proof)
    {
        if (success.FinalRequests.Count != 1)
        {
            throw new InvalidOperationException("Resume response has an unexpected final auth-module set.");
        }

        var module = success.FinalRequests[0];
        if (!IsModule(module.Id, SessionProofModuleIdentity))
        {
            throw new InvalidOperationException("Resume response used an unknown final auth module.");
        }

        if (module.Data.Length != 33 || module.Data[0] != 2)
        {
            throw new InvalidOperationException("Resume response has an invalid session-proof phase.");
        }

        if (!CryptographicOperations.FixedTimeEquals(module.Data.AsSpan(1), proof.ExpectedServerProof))
        {
            throw new InvalidOperationException("Native server session proof failed.");
        }
    }

    private static bool IsModule(ModuleId id, byte[] identity) =>
        id.Usage.AsSpan().SequenceEqual(ModuleUsage) && id.Identity.AsSpan().SequenceEqual(identity);

    public static ResumeResult DecodeResumeResponse(BitReader reader)
    {
        // Resume base struct is 0 fields / 0 bits — the selector is the first bit on the wire.
        var isFailure = reader.Read(1) != 0;
        return isFailure ? DecodeFailure(reader) : DecodeSuccess(reader);
    }

    private static ResumeResult.Success DecodeSuccess(BitReader reader)
    {
        var finalRequestCount = (int)reader.Read(3);
        var finalRequests = new List<ModuleInput>(finalRequestCount);
        for (var i = 0; i < finalRequestCount; i++)
        {
            finalRequests.Add(DecodeModuleInput(reader));
        }

        var pingTimeoutSeconds = ReadS32(reader);

        var regulatorRules = reader.Read(1) != 0 ? DecodeRegulatorInfo(reader) : null;

        return new ResumeResult.Success(finalRequests, pingTimeoutSeconds, regulatorRules);
    }

    /// <summary>
    /// Decodes a Conn/11 RegulatorUpdate's payload — a bare (non-optional)
    /// Battlenet::Regulator::Info choice. Per the docs, this record "can
    /// occur before ResumeResponse"; callers in the middle of Resume should
    /// decode and discard it, then keep waiting for the actual response.
    /// </summary>
    public static RegulatorInfo DecodeRegulatorUpdate(BitReader reader) => DecodeRegulatorInfo(reader);

    private static RegulatorInfo DecodeRegulatorInfo(BitReader reader) => reader.Read(1) != 0
        ? new RegulatorInfo.LeakyBucket((uint)reader.Read(32), (uint)reader.Read(32))
        : new RegulatorInfo.None();

    private static ResumeResult.Failure DecodeFailure(BitReader reader)
    {
        byte[]? strings = reader.Read(1) != 0 ? reader.ReadBytes(40, aligned: true) : null;

        var reasonSelector = reader.Read(2);
        ResumeFailureReason reason = reasonSelector switch
        {
            0 => new ResumeFailureReason.Update(),
            1 => new ResumeFailureReason.Failed((ushort)reader.Read(16), ReadS32(reader)),
            2 => new ResumeFailureReason.VersionCheckDisconnect(),
            _ => throw new InvalidOperationException("Resume failure reason has an unknown choice."),
        };

        return new ResumeResult.Failure(strings, reason);
    }

    private static ModuleInput DecodeModuleInput(BitReader reader)
    {
        var usage = reader.ReadBytes(8, aligned: true);
        var identity = reader.ReadBytes(32, aligned: true);
        var dataLength = (int)reader.Read(10);
        var data = reader.ReadBytes(dataLength, aligned: true);
        return new ModuleInput(new ModuleId(usage, identity), data);
    }

    /// <summary>
    /// Decodes a Battlenet::s32 (core/src/native/schema/wire.rs type 37:
    /// <c>IntegerRange { bit_width: 32, minimum: -2147483648, maximum:
    /// 2147483647 }</c>). BSN's codec always computes <c>value = raw +
    /// minimum</c>; for this type that is NOT the same as reinterpreting the
    /// raw 32 bits as a two's-complement int (a naive <c>unchecked((int)raw)</c>)
    /// — since minimum is exactly -2^31, `raw + minimum` mod 2^32 is equal to
    /// flipping the sign bit before that reinterpretation.
    /// </summary>
    private static int ReadS32(BitReader reader) => unchecked((int)(reader.Read(32) ^ 0x8000_0000UL));

    private static void WriteRequestCommon(BitWriter writer)
    {
        writer.Write(FourCc.Encode("S2"), 32);
        writer.Write(FourCc.Encode("Mc64"), 32);
        writer.Write(FourCc.Encode("enUS"), 32);
        writer.Write((ulong)MacosNativeVersions.Length, 6);
        foreach (var (program, component, version) in MacosNativeVersions)
        {
            writer.Write(FourCc.Encode(program), 32);
            writer.Write(FourCc.Encode(component), 32);
            writer.Write(version, 32);
        }
    }

    /// <summary>Writes a length-prefixed blob whose length field is BSN-biased by <paramref name="minimum"/> (i.e. the wire stores <c>bytes.Length - minimum</c>, per the field's schema range) — see the class remarks for why this matters.</summary>
    private static void WriteBlob(BitWriter writer, byte[] bytes, int lengthBits, int minimum = 0)
    {
        writer.Write((ulong)(bytes.Length - minimum), lengthBits);
        writer.WriteBytes(bytes, aligned: true);
    }
}
