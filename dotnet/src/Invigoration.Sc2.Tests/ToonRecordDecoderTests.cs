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
}
