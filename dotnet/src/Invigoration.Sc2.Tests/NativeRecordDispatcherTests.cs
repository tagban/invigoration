using Invigoration.Sc2.Native;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Tests;

/// <summary>Routing coverage for <see cref="NativeRecordDispatcher"/> — reuses the same retail vectors as <see cref="ChatRecordDecoderTests"/>/<see cref="FriendsRecordDecoderTests"/>, but drives them through the full (slot, command) → decoder switch instead of calling a decoder directly.</summary>
public class NativeRecordDispatcherTests
{
    [Fact]
    public void Decode_RoutesChatMessageToTheChatDecoder()
    {
        var packet = Convert.FromHexString(
            "4b0505012f0b04616e796f6f6e6520666f72206d75746174696f6e206f72206f6e65206d697373696f6e3f00");
        var reader = new BitReader(packet);
        var routing = RoutingHeader.Decode(reader);

        var record = NativeRecordDispatcher.Decode(routing.CommandId, routing.ServiceSlot, reader);

        var message = Assert.IsType<NativeChatRecord.Message>(record);
        Assert.Equal("anyoone for mutation or one mission?", message.Value.Body);
    }

    [Fact]
    public void Decode_RoutesToonsOfFriendsToTheFriendsDecoder()
    {
        var packet = Convert.FromHexString(
            "460301010014cc0200000011004563686f657323323935cafebabe7f1884100000000002fe223701");
        var reader = new BitReader(packet);
        var routing = RoutingHeader.Decode(reader);

        var record = NativeRecordDispatcher.Decode(routing.CommandId, routing.ServiceSlot, reader);

        var toons = Assert.IsType<NativeChatRecord.ToonsOfFriends>(record);
        Assert.Equal(50_209_335u, toons.Value.Entries[0].AccountId);
    }

    [Fact]
    public async Task Decode_ViaRecordStream_ProducesTheSameResultAsCallingTheDecoderDirectly()
    {
        var packet = Convert.FromHexString(
            "460301010014cc0200000011004563686f657323323935cafebabe7f1884100000000002fe223701");
        using var stream = new RecordStream(new MemoryStream(packet));
        await stream.FillAsync();

        var completed = stream.TryDecodeRecord(NativeRecordDispatcher.Decode, out var record);

        Assert.True(completed);
        var toons = Assert.IsType<NativeChatRecord.ToonsOfFriends>(record);
        Assert.True(toons.Value.Complete);
    }

    [Fact]
    public void Decode_UnknownRoute_ThrowsRatherThanSilentlySkipping()
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, commandId: 63, serviceSlot: ChatCommands.ToonSlot);
        writer.Align();
        var reader = new BitReader(writer.ToBytes());
        var routing = RoutingHeader.Decode(reader);

        Assert.Throws<InvalidOperationException>(() => NativeRecordDispatcher.Decode(routing.CommandId, routing.ServiceSlot, reader));
    }
}
