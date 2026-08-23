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

    [Fact]
    public void DecodeToonBlockNotify_RetailVector_DecodesAtExactBoundary()
    {
        // Real captured ToonBlockNotify record (Friends slot, command 33): one
        // entry (a toon being removed from the account's block list), no
        // trailing "complete" flag. Field widths cross-checked bit-exact
        // against the extracted retail SC2 schema (types 2724/2725/2729/2732/
        // 1053) via ncarrillo/superiority's inspect_native_record/decode_hex
        // tooling, which independently reports the same region/program id/
        // realm/name/update values this decoder computes.
        var packet = Convert.FromHexString("61030101000a660200000019025472756d70466c6174");
        var reader = new BitReader(packet);
        RoutingHeader.Decode(reader);

        var record = FriendsRecordDecoder.DecodeToonBlockNotify(reader);

        Assert.Null(record.Complete);
        Assert.Single(record.Entries);
        var entry = record.Entries[0];
        Assert.True(entry.IsRemove);
        // The name's first byte is a literal 0x02 (STX) control character, not
        // a display artifact — confirmed byte-for-byte against the retail
        // capture, and the reference tool's own debug print just doesn't
        // render it visibly.
        Assert.Equal(new ToonFullName(1, 10650, 1, (char)2 + "TrumpFl"), entry.Toon);
        Assert.Equal(162, reader.Position);
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
    public void DecodeFriendsList_AccountWithFullNamePresent_ThrowsRatherThanDesync()
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, commandId: 30, serviceSlot: 3);
        writer.Write(0, 1); // complete: absent
        writer.Write(1, 7); // one update
        writer.Write(0, 2); // operation: Add
        writer.Write(1, 2); // container choice: Account
        writer.Write(1u, 32); // m_accountId
        writer.Write(1, 1); // m_fullName: present (unsupported)
        writer.Align();

        var reader = new BitReader(writer.ToBytes());
        RoutingHeader.Decode(reader);

        Assert.Throws<NotSupportedException>(() => FriendsRecordDecoder.DecodeFriendsList(reader));
    }
}
