using Invigoration.Sc2.Native;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Tests;

/// <summary>
/// SC2 portrait lookup ported from ncarrillo/superiority (MIT): the profile read request
/// (protocol.rs), its answer (decode.rs), the PORT value in a profile block
/// (session.rs), the portrait catalog and the rate-limited resolver.
/// </summary>
public class PortraitTests
{
    private static readonly PlayerTarget.ProfileRecordAddress Address = new(0xcafe_babe, 0xdc82_c80e_0000_0000);
    private static readonly DateTimeOffset T0 = new(2026, 9, 24, 12, 0, 0, TimeSpan.Zero);

    // --- request ---

    [Fact]
    public void ProfileReadRequest_MatchesUpstreamRetailLayout()
    {
        // protocol.rs generated_profile_avatar_read_matches_retail_layout: request 24, path [4].
        var bytes = ChatCommands.ProfileReadRequest(24, Address, [4]);

        Assert.Equal("c00600000000000003c85fd757de905901c0000000000104", Convert.ToHexString(bytes).ToLowerInvariant());
    }

    [Fact]
    public void ProfileReadRequest_DefaultsToTheAvatarPath()
    {
        var bytes = ChatCommands.ProfileReadRequest(24, Address);

        Assert.Equal("c00600000000000003c85fd757de905901c0000000000114", Convert.ToHexString(bytes).ToLowerInvariant());
    }

    // --- response ---

    [Fact]
    public void ProfileRead_RetailStartDecodesAtTheExactBoundary()
    {
        // decode.rs retail_profile_read_start_decodes_at_the_exact_boundary: payload starts at bit 3, 98 bits long.
        var reader = new BitReader(Convert.FromHexString("06000000010000c00100000000"), startPosition: 3);

        var record = ProfileRecordDecoder.DecodeProfileRead(reader);

        Assert.Equal(101, reader.Position);
        Assert.Equal(new ProfileReadRecord(0, ProfileReadKind.Start, PacketCount: 1, RecordType: 6145), record);
    }

    [Fact]
    public void ProfileRead_EveryKindRoutesAndStaysInStep()
    {
        var start = AssertDecodesThenMarker(Response(w => { w.Write(0, 2); w.Write(3, 32); w.Write(6145, 32); }, 7));
        Assert.Equal((ProfileReadKind.Start, 3u, 6145u, 7u), (start.Kind, start.PacketCount, start.RecordType, start.RequestId));

        var block = AssertDecodesThenMarker(Response(w => { w.Write(1, 2); w.Write(3, 14); w.WriteBytes([1, 2, 3], aligned: true); }, 8));
        Assert.Equal(ProfileReadKind.Block, block.Kind);
        Assert.Equal(new byte[] { 1, 2, 3 }, block.Block);
        Assert.Equal(8u, block.RequestId);

        var failure = AssertDecodesThenMarker(Response(w => { w.Write(2, 2); w.Write(0x1234, 16); }, 9));
        Assert.Equal((ProfileReadKind.Failure, (ushort)0x1234, 9u), (failure.Kind, failure.FailureCode, failure.RequestId));

        var cache = AssertDecodesThenMarker(Response(w => w.Write(3, 2), 10));
        Assert.Equal((ProfileReadKind.Cache, 10u), (cache.Kind, cache.RequestId));
    }

    // --- block + catalog ---

    [Fact]
    public void Block_ReadsUpstreamWorkedExamples()
    {
        // session.rs decodes_targeted_avatar_profile_value
        Assert.True(Sc2PortraitBlock.TryReadUnlockableId(Convert.FromHexString("0614f0504f525401f15fce1068"), out var first));
        Assert.Equal(Sc2PortraitCatalog.DefaultUnlockableId, first);
        Assert.Equal(Sc2Portrait.Default, Sc2PortraitBlock.PortraitIn(Convert.FromHexString("0614f0504f525401f15fce1068")));

        Assert.True(Sc2PortraitBlock.TryReadUnlockableId(Convert.FromHexString("0614f0504f525401ebae8364"), out var second));
        Assert.Equal(97_993_138u, second);
        Assert.Equal(new Sc2Portrait(0, 15), Sc2PortraitBlock.PortraitIn(Convert.FromHexString("0614f0504f525401ebae8364")));

        Assert.False(Sc2PortraitBlock.TryReadUnlockableId(Convert.FromHexString("0614f049434f4e0100"), out _));
        Assert.Null(Sc2PortraitBlock.PortraitIn(Convert.FromHexString("0614f049434f4e0100")));
    }

    [Fact]
    public void Block_FindsTheKeyAnywhereAndRejectsTruncatedValues()
    {
        Assert.True(Sc2PortraitBlock.TryReadUnlockableId(Convert.FromHexString("aabbcc0614f0504f525401f15fce1068ff"), out var id));
        Assert.Equal(Sc2PortraitCatalog.DefaultUnlockableId, id);
        Assert.True(Sc2PortraitBlock.TryReadUnlockableId(Convert.FromHexString("0614f0504f52540100"), out var zero));
        Assert.Equal(0u, zero);
        Assert.False(Sc2PortraitBlock.TryReadUnlockableId(Convert.FromHexString("0614f0504f525401f15fce"), out _));
        Assert.False(Sc2PortraitBlock.TryReadUnlockableId(Convert.FromHexString("0614f0504f525401"), out _));
    }

    [Fact]
    public void Catalog_HasEveryUpstreamEntry()
    {
        Assert.Equal(573, Sc2PortraitCatalog.Count);
        Assert.True(Sc2PortraitCatalog.TryGet(2_951_153_716, out var fallback));
        Assert.Equal(Sc2Portrait.Default, fallback);
        Assert.True(Sc2PortraitCatalog.TryGet(905_359, out var first));
        Assert.Equal(new Sc2Portrait(5, 8), first);
        Assert.True(Sc2PortraitCatalog.TryGet(4_288_809_372, out var last));
        Assert.Equal(new Sc2Portrait(10, 28), last);
        Assert.False(Sc2PortraitCatalog.TryGet(1, out _));
    }

    // --- presence ---

    [Fact]
    public void Presence_ReadsAvatarAndProfileFieldsByLocalOrMasterId()
    {
        var tracker = new PresenceTracker();
        tracker.Announce(new PresenceFieldsRecord(
        [
            new PresenceFieldSpec(PresenceTracker.FieldAvatar, 0, 4, false, false, false, false),
            new PresenceFieldSpec(PresenceTracker.FieldToonProfile, 0, 12, false, false, false, false),
        ]));
        byte[] data = [0x00, 0x03, 0x00, 0x11, 0xca, 0xfe, 0xba, 0xbe, 0xdc, 0x82, 0xc8, 0x0e, 0, 0, 0, 0];
        Assert.True(tracker.Apply(new PresenceUpdateRecord(7, 8, true, data, [], [PresenceTracker.FieldAvatar, PresenceTracker.FieldToonProfile], [])));

        Assert.Equal(new Sc2Portrait(3, 17), tracker.AvatarFor(7));
        Assert.Equal(new Sc2Portrait(3, 17), tracker.AvatarFor(8));
        Assert.Equal(Address, tracker.ProfileFor(7));
        Assert.Null(tracker.AvatarFor(99));
        Assert.Null(tracker.ProfileFor(99));

        var resolver = new PortraitResolver();
        Assert.False(resolver.EnqueueMember(tracker, 7)); // the presence avatar wins, no read needed
        Assert.Equal(new Sc2Portrait(3, 17), resolver.For(tracker, 7));
    }

    // --- resolver ---

    [Fact]
    public void Resolver_ReadsEachAddressOnceAndCachesMisses()
    {
        var resolver = new PortraitResolver(firstRequestId: 5);
        Assert.True(resolver.Enqueue(Address));
        Assert.False(resolver.Enqueue(Address));

        var request = Assert.Single(resolver.NextRequests(T0));
        Assert.Equal((5u, Address), request);
        Assert.False(resolver.Enqueue(Address)); // in flight

        Assert.True(resolver.Complete(new ProfileReadRecord(5, ProfileReadKind.Cache)));
        Assert.True(resolver.IsResolved(Address));
        Assert.Null(resolver.For(Address));
        Assert.False(resolver.Enqueue(Address)); // a miss is cached too
        Assert.Empty(resolver.NextRequests(T0.AddSeconds(1)));
    }

    [Fact]
    public void Resolver_AssemblesStartAndBlocks()
    {
        var resolver = new PortraitResolver();
        resolver.Enqueue(Address);
        var (id, _) = Assert.Single(resolver.NextRequests(T0));
        var block = Convert.FromHexString("0614f0504f525401ebae8364");

        Assert.False(resolver.Complete(new ProfileReadRecord(id, ProfileReadKind.Start, PacketCount: 2, RecordType: 6145)));
        Assert.False(resolver.Complete(new ProfileReadRecord(id, ProfileReadKind.Block, Block: block[..6])));
        Assert.Null(resolver.For(Address));
        Assert.True(resolver.Complete(new ProfileReadRecord(id, ProfileReadKind.Block, Block: block[6..])));

        Assert.Equal(new Sc2Portrait(0, 15), resolver.For(Address));
        Assert.Equal(0, resolver.InFlight);
        Assert.False(resolver.Complete(new ProfileReadRecord(id, ProfileReadKind.Block, Block: block))); // late answers are ignored
    }

    [Fact]
    public void Resolver_MissesOnEmptyStartFailureAndBlocksWithoutAPortrait()
    {
        var resolver = new PortraitResolver();
        var addresses = Enumerable.Range(1, 4).Select(i => new PlayerTarget.ProfileRecordAddress(0xcafe_babe, (ulong)i)).ToList();
        addresses.ForEach(a => resolver.Enqueue(a));
        var ids = resolver.NextRequests(T0).Select(r => r.RequestId).ToList();

        Assert.True(resolver.Complete(new ProfileReadRecord(ids[0], ProfileReadKind.Start, PacketCount: 0)));
        Assert.True(resolver.Complete(new ProfileReadRecord(ids[1], ProfileReadKind.Failure, FailureCode: 1)));
        Assert.True(resolver.Complete(new ProfileReadRecord(ids[2], ProfileReadKind.Block, Block: Convert.FromHexString("0614f049434f4e0100"))));
        Assert.False(resolver.Complete(new ProfileReadRecord(ids[3], ProfileReadKind.Start, PacketCount: 1)));
        Assert.True(resolver.Complete(new ProfileReadRecord(ids[3], ProfileReadKind.Block, Block: [1, 2, 3])));

        Assert.All(addresses, a => Assert.True(resolver.IsResolved(a)));
        Assert.All(addresses, a => Assert.Null(resolver.For(a)));
    }

    [Fact]
    public void Resolver_CapsInFlightAndRate()
    {
        var resolver = new PortraitResolver();
        for (var i = 0; i < 40; i++)
        {
            resolver.Enqueue(new PlayerTarget.ProfileRecordAddress(0xcafe_babe, (ulong)i));
        }

        const int inFlight = PortraitResolver.MaxInFlight;
        var burst = (int)PortraitResolver.Burst;
        var perTenthOfASecond = (int)(PortraitResolver.RatePerSecond / 10);

        var first = resolver.NextRequests(T0).ToList();
        Assert.Equal(inFlight, first.Count); // a full bucket, but only so many slots
        Assert.Empty(resolver.NextRequests(T0.AddSeconds(5)));

        foreach (var (id, _) in first)
        {
            resolver.Complete(new ProfileReadRecord(id, ProfileReadKind.Cache));
        }

        // The idle 5 s refilled the bucket only to the burst size; slots still cap it.
        Assert.Equal(inFlight, resolver.NextRequests(T0.AddSeconds(5)).Count());
        foreach (var id in Enumerable.Range(inFlight, inFlight))
        {
            resolver.Complete(new ProfileReadRecord((uint)id, ProfileReadKind.Cache));
        }

        // Same instant: what's left of the burst.
        Assert.Equal(burst - inFlight, resolver.NextRequests(T0.AddSeconds(5)).Count());
        // 100 ms later: a tenth of a second's worth more.
        Assert.Equal(perTenthOfASecond, resolver.NextRequests(T0.AddTicks(TimeSpan.TicksPerSecond * 51 / 10)).Count());
        Assert.Equal(burst - inFlight + perTenthOfASecond, resolver.InFlight);
    }

    [Fact]
    public void Resolver_GivesUpOnUnansweredRequests()
    {
        var resolver = new PortraitResolver();
        resolver.Enqueue(Address);
        Assert.Single(resolver.NextRequests(T0));

        Assert.Empty(resolver.NextRequests(T0 + PortraitResolver.PendingTimeout));

        Assert.Equal(0, resolver.InFlight);
        Assert.True(resolver.IsResolved(Address));
        Assert.Null(resolver.For(Address));
    }

    private static byte[] Response(Action<BitWriter> result, uint requestId)
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, ChatCommands.ProfileReadCommand, ChatCommands.ProfileSlot);
        result(writer);
        writer.Write(requestId, 32);
        writer.Align();
        return writer.ToBytes();
    }

    /// <summary>Decodes <paramref name="record"/> through a RecordStream with a chat message right behind it, which only decodes if the first record ended on exactly the right byte.</summary>
    private static ProfileReadRecord AssertDecodesThenMarker(byte[] record)
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, 11, ChatCommands.ChatSlot);
        writer.Write(42, 32);
        var bytes = System.Text.Encoding.UTF8.GetBytes("marker");
        writer.Write((ulong)bytes.Length, 10);
        writer.WriteBytes(bytes, aligned: true);
        writer.Write(2, 3);
        writer.Align();
        using var stream = new RecordStream(new MemoryStream([.. record, .. writer.ToBytes()]));
        stream.FillAsync().GetAwaiter().GetResult();

        Assert.True(stream.TryDecodeRecord(NativeRecordDispatcher.Decode, out var first));
        Assert.True(stream.TryDecodeRecord(NativeRecordDispatcher.Decode, out var second));
        Assert.Equal("marker", Assert.IsType<NativeChatRecord.Message>(second).Value.Body);
        return Assert.IsType<NativeChatRecord.ProfileRead>(first).Value;
    }
}
