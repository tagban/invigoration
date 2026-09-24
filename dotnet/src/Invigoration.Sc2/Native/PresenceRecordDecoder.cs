using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

// Ported from ncarrillo/superiority (MIT): core/src/games/sc2/native/decode.rs's
// presence_fields_with_provenance and presence_update_with_provenance.

/// <summary>
/// One presence field definition from a FieldSpecAnnounce. Presence updates name
/// fields only by <see cref="Handle"/>; this says how many bytes each one's value
/// takes, or that its size comes from the update's variable-size list when
/// <see cref="FixedSize"/> is null.
/// </summary>
public sealed record PresenceFieldSpec(
    uint Handle,
    byte TypeId,
    ushort? FixedSize,
    bool ClientOnly,
    bool Writable,
    bool Ephemeral,
    bool ServerOnly);

/// <summary>FieldSpecAnnounce (Presence slot, command 1): the field definitions later presence updates rely on.</summary>
public sealed record PresenceFieldsRecord(IReadOnlyList<PresenceFieldSpec> Fields);

/// <summary>
/// PresenceUpdateNotify (Presence slot, command 0): new values for one presence.
/// <see cref="FieldData"/> is the values of <see cref="Handles"/> back to back, in
/// order; each one's length is its field's fixed size, or the next entry of
/// <see cref="VariableSizes"/>. Feed it to <see cref="PresenceTracker.Apply"/>.
/// </summary>
public sealed record PresenceUpdateRecord(
    uint LocalPresenceId,
    uint MasterPresenceId,
    bool Online,
    byte[] FieldData,
    IReadOnlyList<uint> ClearedHandles,
    IReadOnlyList<uint> Handles,
    IReadOnlyList<ushort> VariableSizes);

/// <summary>
/// Decoders for the two presence records a friends list needs. The bit layouts are
/// the ones the earlier skip-only readers consumed in live sessions; these keep the
/// values instead of discarding them.
/// </summary>
public static class PresenceRecordDecoder
{
    /// <summary>
    /// FieldSpecAnnounce (Presence 1): a 7-bit count, then per field client-only,
    /// writable and ephemeral bits, an "absent" bit followed by a 16-bit fixed size
    /// only when it's 0, a server-only bit, an 8-bit type id and a 32-bit handle.
    /// </summary>
    public static PresenceFieldsRecord DecodeFieldSpecAnnounce(BitReader reader)
    {
        var count = (int)reader.Read(7);
        var fields = new List<PresenceFieldSpec>(count);
        for (var i = 0; i < count; i++)
        {
            var clientOnly = reader.Read(1) != 0;
            var writable = reader.Read(1) != 0;
            var ephemeral = reader.Read(1) != 0;
            ushort? fixedSize = reader.Read(1) == 0 ? (ushort)reader.Read(16) : null;
            var serverOnly = reader.Read(1) != 0;
            var typeId = (byte)reader.Read(8);
            var handle = (uint)reader.Read(32);
            fields.Add(new PresenceFieldSpec(handle, typeId, fixedSize, clientOnly, writable, ephemeral, serverOnly));
        }

        return new PresenceFieldsRecord(fields);
    }

    /// <summary>
    /// PresenceUpdateNotify (Presence 0): a 19-bit layout selector, an inverted
    /// online bit (0 means online), local and master presence ids, an 11-bit byte
    /// count and that many byte-aligned bytes of field data, 11 reserved bits, then
    /// three arrays with 4-bit counts (cleared handles and handles, 32 bits each;
    /// variable sizes, 16 bits each), an optional flag-and-id pair and 8 trailing bits.
    /// </summary>
    public static PresenceUpdateRecord DecodePresenceUpdate(BitReader reader)
    {
        reader.Skip(19);
        var online = reader.Read(1) == 0;
        var localPresenceId = (uint)reader.Read(32);
        var masterPresenceId = (uint)reader.Read(32);
        var fieldData = reader.ReadBlob(11);
        reader.Skip(11);
        var clearedHandles = ReadUInt32Array(reader);
        var handles = ReadUInt32Array(reader);
        var sizeCount = (int)reader.Read(4);
        var variableSizes = new ushort[sizeCount];
        for (var i = 0; i < sizeCount; i++)
        {
            variableSizes[i] = (ushort)reader.Read(16);
        }

        reader.SkipOptional(r => r.Skip(1 + 32));
        reader.Read(8);
        return new PresenceUpdateRecord(localPresenceId, masterPresenceId, online, fieldData, clearedHandles, handles, variableSizes);
    }

    private static uint[] ReadUInt32Array(BitReader reader)
    {
        var values = new uint[(int)reader.Read(4)];
        for (var i = 0; i < values.Length; i++)
        {
            values[i] = (uint)reader.Read(32);
        }

        return values;
    }
}
