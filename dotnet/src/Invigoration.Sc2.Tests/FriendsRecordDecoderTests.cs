using Invigoration.Sc2.Native;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Tests;

/// <summary>
/// The ToonsOfFriendsNotify vector is a retail-captured packet reproduced
/// verbatim from ncarrillo/superiority's core/src/native/decode.rs unit test
/// retail_friend_toon_decodes_generated_order_at_the_exact_boundary (relayed
/// via research agent) — same provenance as ChatRecordDecoderTests's
/// vectors. FriendsListNotify5 has no equivalent retail vector upstream, so
/// its tests below are self-consistency round-trips (encode via BitWriter,
/// decode, compare) rather than independently-verified captures — weaker
/// evidence than a retail vector, but still catches real decoder bugs.
/// </summary>
public class FriendsRecordDecoderTests
{
    [Fact]
    public void DecodeToonsOfFriends_RetailVector_DecodesAtExactBoundary()
    {
        var packet = Convert.FromHexString(
            "460301010014cc0200000011004563686f657323323935cafebabe7f1884100000000002fe223701");
        var reader = new BitReader(packet);
        RoutingHeader.Decode(reader);

        var page = FriendsRecordDecoder.DecodeToonsOfFriends(reader);

        Assert.True(page.Complete);
        Assert.Single(page.Entries);
        var entry = page.Entries[0];
        Assert.Equal(50_209_335u, entry.AccountId);
        Assert.Equal(FourCc.Encode("S2"), entry.ProgramId);
        Assert.Equal(new ToonFullName(1, FourCc.Encode("S2"), 1, "Echoes#295"), entry.ToonName);
        Assert.Equal(new PlayerTarget.ProfileRecordAddress(0xcafe_babe, 0x7f18_8410_0000_0000), entry.Profile);
        Assert.Equal(313, reader.Position);
    }

    /// <summary>
    /// The one real ToonBlockNotify capture (Friends slot, command 33) stops
    /// partway through its toon name: the name's length says 16 bytes, but only
    /// "TrumpFlat" (9) was saved. Its header still reads cleanly as region 1,
    /// program "S2", realm 1, which the old 5-bit, choice-bit-last layout got
    /// wrong (program 10650). A cut-off record must read as "wait for more
    /// bytes", never as a finished, garbled one.
    /// </summary>
    private const string TruncatedToonBlockCapture = "61030101000a660200000019025472756d70466c6174";

    [Fact]
    public async Task DecodeToonBlockNotify_TruncatedRetailCapture_WaitsForTheRestOfTheName()
    {
        using var stream = new RecordStream(new MemoryStream(Convert.FromHexString(TruncatedToonBlockCapture)));
        await stream.FillAsync();

        Assert.False(stream.TryDecodeRecord(NativeRecordDispatcher.Decode, out _));
    }

    [Fact]
    public void DecodeToonBlockNotify_RetailCaptureCompleted_ReadsChoiceBitThenFullName()
    {
        // The real capture, completed with a made-up 7-byte name tail
        // ("#123456", making the 16 bytes its length field promises) and a
        // zero byte for the absent "complete" flag. Only the tail is invented.
        var packet = Convert.FromHexString(TruncatedToonBlockCapture + "23313233343536" + "00");
        var reader = new BitReader(packet);
        RoutingHeader.Decode(reader);

        var record = FriendsRecordDecoder.DecodeToonBlockNotify(reader);

        Assert.Null(record.Complete);
        var entry = Assert.Single(record.Entries);
        Assert.False(entry.IsRemove);
        Assert.Equal(new ToonFullName(1, FourCc.Encode("S2"), 1, "TrumpFlat#123456"), entry.Toon);
        Assert.Equal(233, reader.Position);
    }

    [Fact]
    public void DecodeFriendsList_RoundTripsAnAccountAdd()
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, commandId: 30, serviceSlot: 3);
        writer.Write(0, 1); // complete: absent
        writer.Write(1, 7); // one update
        writer.Write(0, 2); // operation: Add
        writer.Write(1, 2); // container choice: Account
        writer.Write(12345u, 32); // m_accountId
        writer.Write(0, 1); // m_fullName: absent
        writer.Write(1, 1); // display_name: present
        writer.Write((ulong)"Nelson".Length, 7);
        writer.WriteBytes(System.Text.Encoding.UTF8.GetBytes("Nelson"), aligned: true);
        writer.Write(0xdead_beefu, 32); // profile.label
        writer.Write(0x1122_3344_5566_7788uL, 64); // profile.id
        writer.Write(0, 1); // custom message: absent
        writer.Write(0, 1); // note: absent
        writer.Write(0u ^ 0x8000_0000u, 32); // last_online = 0, sign-flip encoded
        writer.Write(0uL, 64); // account_serial
        writer.Write(0u, 32); // game_account_id
        writer.Align();

        var reader = new BitReader(writer.ToBytes());
        RoutingHeader.Decode(reader);
        var list = FriendsRecordDecoder.DecodeFriendsList(reader);

        Assert.Null(list.Complete);
        Assert.Single(list.Updates);
        var update = list.Updates[0];
        Assert.Equal(SocialOperation.Add, update.Operation);
        Assert.Equal(new FriendIdentity.Account(12345u), update.Entry.Identity);
        Assert.Equal("Nelson", update.Entry.DisplayName);
        Assert.Null(update.Entry.FullName);
        Assert.Null(update.Entry.Note);
        Assert.Equal(new PlayerTarget.ProfileRecordAddress(0xdead_beef, 0x1122_3344_5566_7788), update.Entry.Profile);
    }

    [Fact]
    public void DecodeFriendsList_RoundTripsARemoveByAccountId()
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, commandId: 30, serviceSlot: 3);
        writer.Write(1, 1); // complete: present
        writer.Write(1, 1); // complete: true
        writer.Write(1, 7); // one update
        writer.Write(1, 2); // operation: Remove
        writer.Write(0, 1); // identity: account id
        writer.Write(999u, 32);
        writer.Align();

        var reader = new BitReader(writer.ToBytes());
        RoutingHeader.Decode(reader);
        var list = FriendsRecordDecoder.DecodeFriendsList(reader);

        Assert.Equal(true, list.Complete);
        Assert.Single(list.Updates);
        var update = list.Updates[0];
        Assert.Equal(SocialOperation.Remove, update.Operation);
        Assert.Equal(new FriendIdentity.Account(999u), update.Entry.Identity);
        Assert.Null(update.Entry.DisplayName);
    }

    [Fact]
    public void DecodeFriendsList_AccountWithFullName_ReadsBothHalvesAndStaysInStep()
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, commandId: 30, serviceSlot: 3);
        writer.Write(0, 1); // complete: absent
        writer.Write(1, 7); // one update
        writer.Write(0, 2); // operation: Add
        writer.Write(1, 2); // container choice: Account
        writer.Write(1u, 32); // m_accountId
        writer.Write(1, 1); // m_fullName: present
        writer.Write(3, 8);
        writer.WriteBytes("Jim"u8.ToArray(), aligned: true);
        writer.Write(6, 8);
        writer.WriteBytes("Raynor"u8.ToArray(), aligned: true);
        writer.Write(1, 1); // display_name: present
        writer.Write(5, 7);
        writer.WriteBytes("Jimmy"u8.ToArray(), aligned: true);
        writer.Write(0, 32); // m_profile.m_label
        writer.Write(0, 64); // m_profile.m_id
        writer.Write(0, 1); // custom message: absent
        writer.Write(0, 1); // note: absent
        writer.Write(0x8000_0000, 32); // last_online
        writer.Write(0, 64); // account_serial
        writer.Write(0, 32); // game_account_id
        writer.Align();

        var reader = new BitReader(writer.ToBytes());
        RoutingHeader.Decode(reader);

        var list = FriendsRecordDecoder.DecodeFriendsList(reader);

        var entry = Assert.Single(list.Updates).Entry;
        Assert.Equal("Jim Raynor", entry.FullName);
        Assert.Equal("Jimmy", entry.DisplayName);
    }
}
