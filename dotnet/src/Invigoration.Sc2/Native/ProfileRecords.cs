using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

// Ported from ncarrillo/superiority (MIT): core/src/games/sc2/native/decode.rs's
// profile_read_with_provenance, the Battlenet::Profile::ProfileDataResponse shape in
// native/schema/wire.rs, and chat/session.rs's decode_avatar_block / decode_profile_uint.

/// <summary>Which arm of Battlenet::Profile::ProfileDataResponse a profile read answer carries.</summary>
public enum ProfileReadKind
{
    /// <summary>The read will be answered in <see cref="ProfileReadRecord.PacketCount"/> Block records (none when 0).</summary>
    Start = 0,

    /// <summary>One packet of the record's data, in <see cref="ProfileReadRecord.Block"/>.</summary>
    Block = 1,

    /// <summary>A Battle.net error code, in <see cref="ProfileReadRecord.FailureCode"/>.</summary>
    Failure = 2,

    /// <summary>"You already have it": no data follows.</summary>
    Cache = 3,
}

/// <summary>
/// Profile slot 14, command 0 from the server: one answer to a
/// <see cref="ChatCommands.ProfileReadRequest(uint, PlayerTarget.ProfileRecordAddress)"/>,
/// matched to it by <see cref="RequestId"/>. Only the fields of <see cref="Kind"/> are set.
/// </summary>
public sealed record ProfileReadRecord(
    uint RequestId,
    ProfileReadKind Kind,
    uint PacketCount = 0,
    uint RecordType = 0,
    byte[]? Block = null,
    ushort FailureCode = 0);

public static class ProfileRecordDecoder
{
    /// <summary>Battlenet::Profile::Block: a blob of at most this many bytes.</summary>
    public const int MaxBlockBytes = 8192;

    /// <summary>
    /// A 2-bit choice, then Start (u32 packet count, u32 record type), Block (a 14-bit
    /// byte count, then that many byte-aligned bytes), Failure (u16) or Cache (nothing),
    /// then the 32-bit request id.
    /// </summary>
    public static ProfileReadRecord DecodeProfileRead(BitReader reader)
    {
        var kind = (ProfileReadKind)reader.Read(2);
        uint packets = 0, recordType = 0;
        byte[]? block = null;
        ushort failure = 0;
        switch (kind)
        {
            case ProfileReadKind.Start:
                packets = (uint)reader.Read(32);
                recordType = (uint)reader.Read(32);
                break;
            case ProfileReadKind.Block:
                var length = (int)reader.Read(14);
                if (length > MaxBlockBytes)
                {
                    throw new InvalidOperationException($"Profile read block is {length} bytes, more than {MaxBlockBytes}.");
                }

                block = reader.ReadBytes(length, aligned: true);
                break;
            case ProfileReadKind.Failure:
                failure = (ushort)reader.Read(16);
                break;
            case ProfileReadKind.Cache:
                break;
        }

        var requestId = (uint)reader.Read(32);
        return new ProfileReadRecord(requestId, kind, packets, recordType, block, failure);
    }
}

/// <summary>Finds the portrait in the data a profile read at <see cref="ChatCommands.ProfileAvatarPath"/> returns.</summary>
public static class Sc2PortraitBlock
{
    /// <summary>The key that precedes the portrait value: 06 14 f0 "PORT" 01.</summary>
    private static ReadOnlySpan<byte> PortraitKey => [0x06, 0x14, 0xf0, (byte)'P', (byte)'O', (byte)'R', (byte)'T', 0x01];

    /// <summary>
    /// Reads the portrait unlockable id that follows the PORT key. It's a packed number:
    /// a marker byte whose high nibble minus 11 says how many bytes follow (at most 4),
    /// its low nibble then those bytes big-endian, all shifted right one bit. A 0 marker
    /// is the value 0.
    /// </summary>
    public static bool TryReadUnlockableId(ReadOnlySpan<byte> block, out uint id)
    {
        id = 0;
        var at = block.IndexOf(PortraitKey);
        if (at < 0)
        {
            return false;
        }

        var value = block[(at + PortraitKey.Length)..];
        if (value.IsEmpty)
        {
            return false;
        }

        var marker = value[0];
        if (marker == 0)
        {
            return true;
        }

        var byteCount = (marker >> 4) - 11;
        if (byteCount is < 0 or > 4 || value.Length < 1 + byteCount)
        {
            return false;
        }

        ulong encoded = (ulong)(marker & 0x0f);
        foreach (var b in value.Slice(1, byteCount))
        {
            encoded = (encoded << 8) | b;
        }

        encoded >>= 1;
        if (encoded > uint.MaxValue)
        {
            return false;
        }

        id = (uint)encoded;
        return true;
    }

    /// <summary>The portrait a block names, or null when it names none or one not in <see cref="Sc2PortraitCatalog"/>.</summary>
    public static Sc2Portrait? PortraitIn(ReadOnlySpan<byte> block) =>
        TryReadUnlockableId(block, out var id) && Sc2PortraitCatalog.TryGet(id, out var portrait) ? portrait : null;
}
