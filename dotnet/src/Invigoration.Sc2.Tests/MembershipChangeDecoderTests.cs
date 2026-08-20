using Invigoration.Sc2.Native;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Tests;

/// <summary>
/// No retail-captured packet exists for MembershipChangeNotify (unlike
/// MessageRecv/WhisperRecv), so these are synthetic round-trips built
/// directly against the field-level schema from the SC2Docs type registry,
/// not independent verification against a real server. They still catch
/// bit-width and field-order mistakes in the decoder itself.
/// </summary>
public class MembershipChangeDecoderTests
{
    [Fact]
    public void Decode_JoinAndLeave_ParsesBothChanges()
    {
        var writer = new BitWriter();
        writer.Write(1, 1); // endOfInitial
        writer.Write(2, 3); // channelIndex
        writer.Write(1, 6); // changeCount - 1 (2 changes)

        // Change 1: JoinChannel, one Active status
        writer.Write(1, 2); // selector: joinChannel
        writer.Write(12345, 32); // memberHandle
        writer.Write(999, 32); // presenceId
        writer.Write(1, 3); // statusCount
        writer.Write(6, 3); // status selector: Active
        writer.Write(1, 1); // value: true

        // Change 2: LeaveChannel
        writer.Write(0, 2); // selector: leaveChannel
        writer.Write(54321, 32); // memberHandle
        writer.Write(42, 16); // reason
        writer.Align();

        var reader = new BitReader(writer.ToBytes());
        var record = MembershipChangeDecoder.Decode(reader);

        Assert.True(record.EndOfInitial);
        Assert.Equal(2, record.ChannelIndex);
        Assert.Equal(2, record.Changes.Count);

        var join = Assert.IsType<MembershipChange.Join>(record.Changes[0]);
        Assert.Equal(12345u, join.MemberHandle);
        Assert.Equal(999u, join.PresenceId);
        var status = Assert.Single(join.Statuses);
        var active = Assert.IsType<MemberStatus.Active>(status);
        Assert.True(active.Value);

        var leave = Assert.IsType<MembershipChange.Leave>(record.Changes[1]);
        Assert.Equal(54321u, leave.MemberHandle);
        Assert.Equal(42, leave.Reason);
    }

    [Fact]
    public void Decode_UpdateStatusParty_ParsesOptionalExpansionLevel()
    {
        var writer = new BitWriter();
        writer.Write(0, 1); // endOfInitial
        writer.Write(0, 3); // channelIndex
        writer.Write(0, 6); // changeCount - 1 (1 change)

        writer.Write(2, 2); // selector: updateStatus
        writer.Write(777, 32); // memberHandle
        writer.Write(1, 3); // status selector: Party
        writer.Write(1, 2); // partyStatus: ONLINE
        writer.Write(1, 1); // expansionLevel present
        writer.Write(3, 2); // expansionLevel: LEGACY_OF_THE_VOID
        writer.Write(1, 1); // captain: true
        writer.Align();

        var reader = new BitReader(writer.ToBytes());
        var record = MembershipChangeDecoder.Decode(reader);

        var update = Assert.IsType<MembershipChange.Update>(Assert.Single(record.Changes));
        var party = Assert.IsType<MemberStatus.Party>(update.Status);
        Assert.Equal(1, party.PartyStatus);
        Assert.Equal((byte?)3, party.ExpansionLevel);
        Assert.True(party.Captain);
    }

    [Fact]
    public void Decode_UpdateStatusDisplay_ParsesToonFullName()
    {
        var writer = new BitWriter();
        writer.Write(0, 1);
        writer.Write(0, 3);
        writer.Write(0, 6);

        writer.Write(2, 2); // selector: updateStatus
        writer.Write(1, 32); // memberHandle
        writer.Write(5, 3); // status selector: Display
        writer.Write(1, 8); // region
        writer.Write(FourCc.Encode("S2"), 32); // programId
        writer.Write(1, 32); // realm
        var name = "Tagban"u8.ToArray();
        writer.Write((ulong)(name.Length - 2), 5); // byte count, biased -2
        writer.WriteBytes(name, aligned: true);
        writer.Align();

        var reader = new BitReader(writer.ToBytes());
        var record = MembershipChangeDecoder.Decode(reader);

        var update = Assert.IsType<MembershipChange.Update>(Assert.Single(record.Changes));
        var display = Assert.IsType<MemberStatus.Display>(update.Status);
        Assert.Equal(1, display.ToonName.Region);
        Assert.Equal(FourCc.Encode("S2"), display.ToonName.ProgramId);
        Assert.Equal(1u, display.ToonName.Realm);
        Assert.Equal("Tagban", display.ToonName.Name);
    }

    [Fact]
    public void Decode_TalkerInfoWithToonHandle_UsesGeneratedWireOrder()
    {
        var writer = new BitWriter();
        writer.Write(0, 1);
        writer.Write(0, 3);
        writer.Write(0, 6);

        writer.Write(2, 2); // selector: updateStatus
        writer.Write(1, 32); // memberHandle
        writer.Write(3, 3); // status selector: TalkerInfo
        // TalkerId (m_id) decodes BEFORE the trailing enabled bit.
        writer.Write(1, 2); // TalkerId selector: DatagramConnectionEndPoint
        writer.Write(5, 3); // PlayerTarget selector: toonHandle
        writer.Write(FourCc.Encode("S2"), 32); // programId (written first, per generated layout)
        writer.Write(9, 8); // region
        writer.Write(2, 32); // realm
        writer.Write(123456789UL, 64); // id
        writer.Write(1, 1); // trailing m_enabled bit
        writer.Align();

        var reader = new BitReader(writer.ToBytes());
        var record = MembershipChangeDecoder.Decode(reader);

        var update = Assert.IsType<MembershipChange.Update>(Assert.Single(record.Changes));
        var talkerInfo = Assert.IsType<MemberStatus.TalkerInfo>(update.Status);
        Assert.True(talkerInfo.Enabled);
        var endpoint = Assert.IsType<TalkerId.DatagramConnectionEndPoint>(talkerInfo.Id);
        var handle = Assert.IsType<PlayerTarget.ToonHandle>(endpoint.Target);
        Assert.Equal(9, handle.Region);
        Assert.Equal(FourCc.Encode("S2"), handle.ProgramId);
        Assert.Equal(2u, handle.Realm);
        Assert.Equal(123456789UL, handle.Id);
    }
}
