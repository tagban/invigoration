using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

/// <summary>Decoded payload of Toon/6 ToonSelected — confirms the toon chosen by Toon/5 ToonSelect (<see cref="ChatCommands.ToonSelect"/>).</summary>
public sealed record ToonSelectedRecord(uint RecordLabel, ulong RecordId, ToonHandleValue Handle, uint Realm, int LastLogon, string ToonName);

/// <summary>Battlenet::Toon::Handle — note the wire order (program, region, realm, id) does NOT match its type-registry field declaration order; confirmed independently from both the SC2Docs service-endpoint listing and the golden-vector-verified chat_whisper native builder.</summary>
public sealed record ToonHandleValue(uint Program, byte Region, uint Realm, ulong Id);

/// <summary>One entry in a ToonList (Toon/0 ToonList). Mirrors core/src/native/model.rs's ToonDisplay — the record address, last-online timestamp, and obfuscation-selector/flags fields that precede <see cref="Realm"/> on the wire are read but not carried into this record, matching upstream's own ToonDisplay shape.</summary>
public sealed record ToonDisplayRecord(string Name, uint Realm);

/// <summary>Decoded payload of a ToonList record (Toon slot, command 0) — the roster shown to pick a character from after logging in, before any toon is selected. Mirrors core/src/native/model.rs's ToonList.</summary>
public sealed record ToonListRecord(IReadOnlyList<ToonDisplayRecord> Displays);

public static class ToonRecordDecoder
{
    /// <summary>
    /// Decodes Toon/0 ToonList. Ported from core/src/native/decode.rs's
    /// toon_list_with_provenance. Unlike <see cref="DecodeToonSelected"/>,
    /// this is NOT independently retail-vector-verified — upstream's own
    /// test for it (toon_list_traces_every_physical_wire_field) is
    /// synthetic, round-tripped through its own generic schema-driven codec
    /// for the m_profile/m_realm fields rather than checked against a
    /// captured packet. The 32-bit width used here for both matches every
    /// other confirmed use of those two field kinds elsewhere in this
    /// project (ProfileRecordAddress's label, and every realm field in
    /// ToonHandle/ToonFullName/ToonSelected) — high confidence by
    /// consistency, but flagged here as the one part of this decoder that
    /// isn't independently nailed down.
    /// </summary>
    public static ToonListRecord DecodeToonList(BitReader reader)
    {
        var count = (int)reader.Read(6);
        if (count > 50)
        {
            throw new InvalidOperationException("Native ToonList contains too many displays.");
        }

        var displays = new List<ToonDisplayRecord>(count);
        for (var i = 0; i < count; i++)
        {
            var byteCount = (int)reader.Read(7) + 2;
            var nameBytes = reader.ReadBytes(byteCount, aligned: true);
            var name = System.Text.Encoding.UTF8.GetString(nameBytes);
            reader.Read(32); // last_online (sign-flip int32), discarded — not carried on ToonDisplay.
            reader.Read(3); // wire_layout_selector, discarded.
            reader.Read(32); // flags, discarded.
            reader.Read(32); // m_profile.m_label, discarded.
            reader.Read(64); // m_profile.m_id, discarded.
            var realm = (uint)reader.Read(32);
            displays.Add(new ToonDisplayRecord(name, realm));
        }

        return new ToonListRecord(displays);
    }

    /// <summary>Decodes Toon/6 ToonSelected. Field order here is service-endpoint "generated order" (true wire order), not the type-registry declaration order.</summary>
    public static ToonSelectedRecord DecodeToonSelected(BitReader reader)
    {
        var recordLabel = (uint)reader.Read(32);
        var recordId = reader.Read(64);
        var handle = DecodeToonHandle(reader);
        var realm = (uint)reader.Read(32);
        var lastLogon = unchecked((int)reader.Read(32));
        var byteCount = (int)reader.Read(7) + 2;
        var nameBytes = reader.ReadBytes(byteCount, aligned: true);
        var toonName = System.Text.Encoding.UTF8.GetString(nameBytes);
        return new ToonSelectedRecord(recordLabel, recordId, handle, realm, lastLogon, toonName);
    }

    private static ToonHandleValue DecodeToonHandle(BitReader reader)
    {
        var program = (uint)reader.Read(32);
        var region = (byte)reader.Read(8);
        var realm = (uint)reader.Read(32);
        var id = reader.Read(64);
        return new ToonHandleValue(program, region, realm, id);
    }
}
