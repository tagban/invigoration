using System.Buffers.Binary;
using Invigoration.Sc2.Native;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Tests;

/// <summary>
/// Presence records built bit by bit from their layouts (no captures yet), each
/// followed by a marker record so a decoder that reads a bit too many or too few
/// fails, plus <see cref="PresenceTracker"/> behaviour ported from
/// ncarrillo/superiority's presence.rs.
/// </summary>
public class PresenceTests
{
    private const uint SessionMark = 0x0001_0003;
    private const uint ClanTag = 0x50004;

    [Fact]
    public void FieldSpecAnnounce_DecodesEveryFieldAndStaysInStep()
    {
        var record = Record(4, 1, w =>
        {
            w.Write(2, 7);
            // client-only, writable, ephemeral; fixed size 4; server-only; type 21; handle
            w.Write(1, 1); w.Write(0, 1); w.Write(1, 1);
            w.Write(0, 1); w.Write(4, 16);
            w.Write(1, 1); w.Write(21, 8); w.Write(PresenceTracker.FieldAccountId, 32);
            // no fixed size
            w.Write(0, 1); w.Write(1, 1); w.Write(0, 1);
            w.Write(1, 1);
            w.Write(0, 1); w.Write(4, 8); w.Write(ClanTag, 32);
        });

        var fields = AssertDecodesThenMarker<NativeChatRecord.PresenceFields>(record).Value.Fields;

        Assert.Equal(new PresenceFieldSpec(PresenceTracker.FieldAccountId, 21, 4, ClientOnly: true, Writable: false, Ephemeral: true, ServerOnly: true), fields[0]);
        Assert.Equal(new PresenceFieldSpec(ClanTag, 4, null, ClientOnly: false, Writable: true, Ephemeral: false, ServerOnly: false), fields[1]);
    }

    [Fact]
    public void PresenceUpdate_DecodesEveryValueAndStaysInStep()
    {
        var record = UpdateRecord(
            local: 0x1111_2222,
            master: 0x3333_4444,
            online: true,
            data: [0, 0, 0, 42, 1],
            cleared: [PresenceTracker.FieldAway],
            handles: [PresenceTracker.FieldAccountId, SessionMark],
            sizes: [7],
            optionalTarget: true);

        var update = AssertDecodesThenMarker<NativeChatRecord.PresenceUpdate>(record).Value;

        Assert.Equal(0x1111_2222u, update.LocalPresenceId);
        Assert.Equal(0x3333_4444u, update.MasterPresenceId);
        Assert.True(update.Online);
        Assert.Equal(new byte[] { 0, 0, 0, 42, 1 }, update.FieldData);
        Assert.Equal(new[] { PresenceTracker.FieldAway }, update.ClearedHandles);
        Assert.Equal(new[] { PresenceTracker.FieldAccountId, SessionMark }, update.Handles);
        Assert.Equal(new ushort[] { 7 }, update.VariableSizes);
    }

    [Fact]
    public void PresenceUpdate_OnlineBitIsInverted()
    {
        var record = UpdateRecord(1, 0, online: false, data: [], cleared: [], handles: [], sizes: []);

        Assert.False(AssertDecodesThenMarker<NativeChatRecord.PresenceUpdate>(record).Value.Online);
    }

    [Fact]
    public void Tracker_DecodedRecords_ShowAFriendOnline()
    {
        var tracker = new PresenceTracker();
        var fields = DecodeOne(Record(4, 1, w =>
        {
            w.Write(2, 7);
            WriteField(w, PresenceTracker.FieldAccountId, 4, typeId: 21);
            WriteField(w, SessionMark, 1, typeId: 0);
        }));
        var update = DecodeOne(UpdateRecord(7, 8, online: true, data: [.. Account(42), 1], cleared: [], handles: [PresenceTracker.FieldAccountId, SessionMark], sizes: []));

        tracker.Announce(Assert.IsType<NativeChatRecord.PresenceFields>(fields).Value);
        Assert.True(tracker.Apply(Assert.IsType<NativeChatRecord.PresenceUpdate>(update).Value));

        Assert.Equal(FriendPresence.Online, tracker.For(new FriendIdentity.Account(42)));
        Assert.Equal(7u, tracker.PresenceIdFor(new FriendIdentity.Account(42)));
    }

    [Fact]
    public void Tracker_UnknownFriend_IsNull()
    {
        var tracker = Announced();
        tracker.Apply(Update(7, 8, true, (PresenceTracker.FieldAccountId, Account(42))));

        Assert.Null(tracker.For(new FriendIdentity.Account(43)));
    }

    [Fact]
    public void Tracker_AwayBusyAndInGame()
    {
        var tracker = Announced();
        var friend = new FriendIdentity.Account(42);
        tracker.Apply(Update(7, 8, true, (PresenceTracker.FieldAccountId, Account(42)), (SessionMark, [1])));

        tracker.Apply(Update(7, 8, true, (PresenceTracker.FieldAway, [1])));
        Assert.Equal(FriendPresence.Away, tracker.For(friend));

        // The most recently written away/busy field wins.
        tracker.Apply(Update(7, 8, true, (PresenceTracker.FieldBusy, [1])));
        Assert.Equal(FriendPresence.Busy, tracker.For(friend));

        tracker.Apply(Update(7, 8, true, (PresenceTracker.FieldInGame, [1])));
        Assert.Equal(FriendPresence.InGame, tracker.For(friend));

        tracker.Apply(Update(7, 8, true, (PresenceTracker.FieldInGame, [0]), (PresenceTracker.FieldBusy, [0])));
        Assert.Equal(FriendPresence.Online, tracker.For(friend));
    }

    [Fact]
    public void Tracker_ClearedField_NoLongerCounts()
    {
        var tracker = Announced();
        var friend = new FriendIdentity.Account(42);
        tracker.Apply(Update(7, 8, true, (PresenceTracker.FieldAccountId, Account(42)), (PresenceTracker.FieldAway, [1])));
        Assert.Equal(FriendPresence.Away, tracker.For(friend));

        tracker.Apply(new PresenceUpdateRecord(7, 8, true, [], [PresenceTracker.FieldAway], [], []));

        Assert.Equal(FriendPresence.Online, tracker.For(friend));
    }

    [Fact]
    public void Tracker_OfflineBit_AndLosingTheSessionMarks_BothMeanOffline()
    {
        var tracker = Announced();
        var friend = new FriendIdentity.Account(42);
        tracker.Apply(Update(7, 8, true, (PresenceTracker.FieldAccountId, Account(42)), (SessionMark, [1])));
        Assert.Equal(FriendPresence.Online, tracker.For(friend));

        tracker.Apply(new PresenceUpdateRecord(7, 8, true, [], [SessionMark], [], []));
        Assert.Equal(FriendPresence.Offline, tracker.For(friend));

        tracker.Apply(Update(7, 8, true, (SessionMark, [1])));
        Assert.Equal(FriendPresence.Online, tracker.For(friend));

        tracker.Apply(Update(7, 8, false));
        Assert.Equal(FriendPresence.Offline, tracker.For(friend));
    }

    [Fact]
    public void Tracker_LinksACharacterByToonHandle()
    {
        var tracker = Announced();
        var handle = new byte[17];
        handle[0] = 1;
        BinaryPrimitives.WriteUInt32BigEndian(handle.AsSpan(1), FourCc.Encode("S2"));
        BinaryPrimitives.WriteUInt32BigEndian(handle.AsSpan(5), 1);
        BinaryPrimitives.WriteUInt64BigEndian(handle.AsSpan(9), 123_456);
        tracker.Apply(Update(7, 8, true, (PresenceTracker.FieldToonHandle, handle), (PresenceTracker.FieldBusy, [1])));

        Assert.Equal(FriendPresence.Busy, tracker.For(new FriendIdentity.Character(FourCc.Encode("S2"), 1, 1, 123_456)));
        Assert.Null(tracker.For(new FriendIdentity.Character(FourCc.Encode("S2"), 1, 1, 999)));
    }

    [Fact]
    public void Tracker_FallsBackToTheFriendsProfileAddress()
    {
        var tracker = Announced();
        var profile = new byte[12];
        BinaryPrimitives.WriteUInt32BigEndian(profile, 0xAABB_CCDD);
        BinaryPrimitives.WriteUInt64BigEndian(profile.AsSpan(4), 77);
        tracker.Apply(Update(7, 8, true, (PresenceTracker.FieldToonProfile, profile)));
        var friend = new FriendEntry(new FriendIdentity.Account(42), "Friend#1234", null, null, new PlayerTarget.ProfileRecordAddress(0xAABB_CCDD, 77), null);

        Assert.Null(tracker.For(friend.Identity));
        Assert.Equal(FriendPresence.Online, tracker.For(friend));
        Assert.Equal(7u, tracker.PresenceIdFor(friend));
    }

    [Fact]
    public void Tracker_VariableSizeFieldsAreSplitCorrectly()
    {
        var tracker = Announced();
        var update = new PresenceUpdateRecord(7, 8, true, [3, 0, (byte)'a', (byte)'b', (byte)'c', (byte)'d', (byte)'e', (byte)'f', 0, 0, 0, 42, 1],
            [], [ClanTag, PresenceTracker.FieldAccountId, PresenceTracker.FieldAway], [8]);

        Assert.True(tracker.Apply(update));
        Assert.Equal(FriendPresence.Away, tracker.For(new FriendIdentity.Account(42)));
    }

    [Fact]
    public void Tracker_MalformedUpdate_IsRejectedAndChangesNothing()
    {
        var tracker = Announced();
        tracker.Apply(Update(7, 8, true, (PresenceTracker.FieldAccountId, Account(42))));

        // Fixed-size away field, but no data for it.
        Assert.False(tracker.Apply(new PresenceUpdateRecord(7, 8, false, [], [], [PresenceTracker.FieldAway], [])));
        // Data the fields don't account for.
        Assert.False(tracker.Apply(new PresenceUpdateRecord(7, 8, false, [1, 2], [], [PresenceTracker.FieldAway], [])));
        // No presence id at all.
        Assert.False(tracker.Apply(new PresenceUpdateRecord(0, 0, false, [], [], [], [])));

        Assert.Equal(FriendPresence.Online, tracker.For(new FriendIdentity.Account(42)));
    }

    [Fact]
    public void Tracker_UnannouncedField_KeepsTheValuesBeforeIt()
    {
        var tracker = Announced();
        var update = new PresenceUpdateRecord(7, 8, true, [0, 0, 0, 42, 9, 9], [], [PresenceTracker.FieldAccountId, 0xDEAD], []);

        Assert.True(tracker.Apply(update));
        Assert.Equal(FriendPresence.Online, tracker.For(new FriendIdentity.Account(42)));
    }

    [Fact]
    public void Tracker_UpdatesByMasterIdReachTheSamePresence()
    {
        var tracker = Announced();
        var friend = new FriendIdentity.Account(42);
        tracker.Apply(Update(7, 8, true, (PresenceTracker.FieldAccountId, Account(42))));

        tracker.Apply(Update(0, 8, true, (PresenceTracker.FieldAway, [1])));

        Assert.Equal(FriendPresence.Away, tracker.For(friend));
        Assert.Equal(7u, tracker.PresenceIdFor(friend));
    }

    [Fact]
    public void Tracker_TwoPresencesForOneAccount_FoldIntoTheNewestOne()
    {
        var tracker = Announced();
        var friend = new FriendIdentity.Account(42);
        tracker.Apply(Update(7, 8, true, (PresenceTracker.FieldAccountId, Account(42)), (PresenceTracker.FieldAway, [1])));
        tracker.Apply(Update(17, 18, true, (PresenceTracker.FieldAccountId, Account(42)), (PresenceTracker.FieldInGame, [1])));

        Assert.Equal(FriendPresence.InGame, tracker.For(friend));
        Assert.Equal(FriendPresence.InGame, tracker.State(7));
    }

    [Fact]
    public void Tracker_SocialAccountIdOutranksThePlainOne()
    {
        var tracker = Announced();
        tracker.Apply(Update(7, 8, true, (PresenceTracker.FieldSocialAccountId, Account(42)), (PresenceTracker.FieldAccountId, Account(99))));

        Assert.Equal(FriendPresence.Online, tracker.For(new FriendIdentity.Account(42)));
        Assert.Null(tracker.For(new FriendIdentity.Account(99)));
    }

    /// <summary>A tracker with every field these tests use announced.</summary>
    private static PresenceTracker Announced()
    {
        var tracker = new PresenceTracker();
        tracker.Announce(new PresenceFieldsRecord(
        [
            Field(PresenceTracker.FieldAccountId, 4, 21),
            Field(PresenceTracker.FieldSocialAccountId, 4, 21),
            Field(SessionMark, 1),
            Field(PresenceTracker.FieldAway, 1),
            Field(PresenceTracker.FieldBusy, 1),
            Field(PresenceTracker.FieldInGame, 1),
            Field(PresenceTracker.FieldToonHandle, 17),
            Field(PresenceTracker.FieldToonProfile, 12),
            Field(ClanTag, null),
        ]));
        return tracker;
    }

    private static PresenceFieldSpec Field(uint handle, ushort? fixedSize, byte typeId = 0) =>
        new(handle, typeId, fixedSize, false, false, false, false);

    /// <summary>An update carrying fixed-size fields only.</summary>
    private static PresenceUpdateRecord Update(uint local, uint master, bool online, params (uint Handle, byte[] Value)[] values) =>
        new(local, master, online, [.. values.SelectMany(v => v.Value)], [], [.. values.Select(v => v.Handle)], []);

    private static byte[] Account(uint accountId)
    {
        var bytes = new byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(bytes, accountId);
        return bytes;
    }

    private static void WriteField(BitWriter w, uint handle, ushort fixedSize, byte typeId)
    {
        w.Write(0, 3);
        w.Write(0, 1); w.Write(fixedSize, 16);
        w.Write(0, 1); w.Write(typeId, 8); w.Write(handle, 32);
    }

    private static byte[] UpdateRecord(uint local, uint master, bool online, byte[] data, uint[] cleared, uint[] handles, ushort[] sizes, bool optionalTarget = false) =>
        Record(4, 0, w =>
        {
            w.Write(0x5A5A5, 19);
            w.Write(online ? 0UL : 1UL, 1);
            w.Write(local, 32);
            w.Write(master, 32);
            w.Write((ulong)data.Length, 11);
            w.WriteBytes(data, aligned: true);
            w.Write(0, 11);
            w.Write((ulong)cleared.Length, 4);
            foreach (var handle in cleared)
            {
                w.Write(handle, 32);
            }

            w.Write((ulong)handles.Length, 4);
            foreach (var handle in handles)
            {
                w.Write(handle, 32);
            }

            w.Write((ulong)sizes.Length, 4);
            foreach (var size in sizes)
            {
                w.Write(size, 16);
            }

            w.Write(optionalTarget ? 1UL : 0UL, 1);
            if (optionalTarget)
            {
                w.Write(1, 1);
                w.Write(0xCAFE_BABE, 32);
            }

            w.Write(0xA5, 8);
        });

    private static byte[] Record(byte slot, byte command, Action<BitWriter> payload)
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, command, slot);
        payload(writer);
        writer.Align();
        return writer.ToBytes();
    }

    private static NativeChatRecord DecodeOne(byte[] record)
    {
        var reader = new BitReader(record);
        var routing = RoutingHeader.Decode(reader);
        return NativeRecordDispatcher.Decode(routing.CommandId, routing.ServiceSlot, reader);
    }

    /// <summary>Decodes <paramref name="record"/> through a RecordStream with a chat message right behind it, which only decodes if the first record ended on exactly the right byte.</summary>
    private static T AssertDecodesThenMarker<T>(byte[] record)
        where T : NativeChatRecord
    {
        var marker = Record(ChatCommands.ChatSlot, 11, w =>
        {
            w.Write(42, 32);
            var bytes = System.Text.Encoding.UTF8.GetBytes("marker");
            w.Write((ulong)bytes.Length, 10);
            w.WriteBytes(bytes, aligned: true);
            w.Write(2, 3);
        });
        using var stream = new RecordStream(new MemoryStream([.. record, .. marker]));
        stream.FillAsync().GetAwaiter().GetResult();

        Assert.True(stream.TryDecodeRecord(NativeRecordDispatcher.Decode, out var first));
        Assert.True(stream.TryDecodeRecord(NativeRecordDispatcher.Decode, out var second));
        Assert.Equal("marker", Assert.IsType<NativeChatRecord.Message>(second).Value.Body);
        return Assert.IsType<T>(first);
    }
}
