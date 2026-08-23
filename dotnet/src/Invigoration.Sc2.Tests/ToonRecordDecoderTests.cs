using Invigoration.Sc2.Native;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Tests;

/// <summary>
/// No captured packet exists for ToonSelected, so this is a synthetic
/// round-trip built against the SC2Docs service-endpoint schema (which,
/// unlike the type registry, presents true wire order) rather than
/// independent verification against a real server.
/// </summary>
public class ToonRecordDecoderTests
{
    [Fact]
    public void DecodeToonSelected_ParsesAllFields()
    {
        var writer = new BitWriter();
        writer.Write(1001, 32); // record_address.label
        writer.Write(2002, 64); // record_address.id
        writer.Write(FourCc.Encode("S2"), 32); // toon_handle.program
        writer.Write(1, 8); // toon_handle.region
        writer.Write(3, 32); // toon_handle.realm
        writer.Write(999999UL, 64); // toon_handle.id
        writer.Write(3, 32); // realm
        writer.Write((uint)1700000000, 32); // last_logon
        var name = "Tagban"u8.ToArray();
        writer.Write((ulong)(name.Length - 2), 7);
        writer.WriteBytes(name, aligned: true);
        writer.Align();

        var record = ToonRecordDecoder.DecodeToonSelected(new BitReader(writer.ToBytes()));

        Assert.Equal(1001u, record.RecordLabel);
        Assert.Equal(2002ul, record.RecordId);
        Assert.Equal(FourCc.Encode("S2"), record.Handle.Program);
        Assert.Equal(1, record.Handle.Region);
        Assert.Equal(3u, record.Handle.Realm);
        Assert.Equal(999999ul, record.Handle.Id);
        Assert.Equal(3u, record.Realm);
        Assert.Equal(1700000000, record.LastLogon);
        Assert.Equal("Tagban", record.ToonName);
    }

    /// <summary>Same caveat as DecodeToonSelected_ParsesAllFields, plus: upstream's own test for this record is itself synthetic (round-tripped through its generic schema codec for m_profile/m_realm), not a captured packet — see DecodeToonList's remarks.</summary>
    [Fact]
    public void DecodeToonList_ParsesEveryDisplay()
    {
        var writer = new BitWriter();
        writer.Write(2, 6); // 2 displays

        var first = "Nova!"u8.ToArray();
        writer.Write((ulong)(first.Length - 2), 7);
        writer.WriteBytes(first, aligned: true);
        writer.Write((uint)123 ^ 0x8000_0000u, 32); // last_online (sign-flip encoded)
        writer.Write(5, 3); // wire_layout_selector
        writer.Write(0xa1b2_c3d4u, 32); // flags
        writer.Write(0x1020_3040u, 32); // profile.label
        writer.Write(0x1122_3344_5566_7788uL, 64); // profile.id
        writer.Write(7u, 32); // realm

        var second = "Raynor"u8.ToArray();
        writer.Write((ulong)(second.Length - 2), 7);
        writer.WriteBytes(second, aligned: true);
        writer.Write(0u ^ 0x8000_0000u, 32);
        writer.Write(0, 3);
        writer.Write(0u, 32);
        writer.Write(0u, 32);
        writer.Write(0uL, 64);
        writer.Write(1u, 32); // realm
        writer.Align();

        var list = ToonRecordDecoder.DecodeToonList(new BitReader(writer.ToBytes()));

        Assert.Equal(2, list.Displays.Count);
        Assert.Equal("Nova!", list.Displays[0].Name);
        Assert.Equal(7u, list.Displays[0].Realm);
        Assert.Equal("Raynor", list.Displays[1].Name);
        Assert.Equal(1u, list.Displays[1].Realm);
    }
}
