using Invigoration.Sc2.Native;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Tests;

/// <summary>
/// Keepalives and the records read only to stay in step with the stream. There
/// are no captures for these yet, so each record is built with
/// <see cref="BitWriter"/> from its documented layout, followed by a marker
/// record: if a decoder reads one bit too many or too few, the marker won't
/// decode. The SC2 chat byte examples are the published ones, not built here.
/// </summary>
public class SessionRecordTests
{
    [Fact]
    public void ChatMessage_MatchesThePublishedByteExample() =>
        Assert.Equal("4b050568656c6c6f00", Convert.ToHexString(ChatCommands.ChatMessage(0, "hello")).ToLowerInvariant());

    [Fact]
    public void ChatJoinPublic_MatchesThePublishedByteExample() =>
        Assert.Equal("40752b72aa13200900000001", Convert.ToHexString(ChatCommands.ChatJoinPublic(1033, 1, "enUS")).ToLowerInvariant());

    [Fact]
    public void ChatWhisper_ToToonName_MatchesThePublishedByteExample()
    {
        var bytes = ChatCommands.ChatWhisper(new Chat.WhisperTarget.ToonName("Test", 1, FourCc.Encode("S2"), 1), "hello");

        Assert.Equal("530d0100014c32000000010254657374010168656c6c6f", Convert.ToHexString(bytes).ToLowerInvariant());
    }

    [Fact]
    public void ChatChannelListRequest_IsJustTheRoute()
    {
        var reader = new BitReader(ChatCommands.ChatChannelListRequest());
        var routing = RoutingHeader.Decode(reader);

        Assert.Equal((byte)21, routing.CommandId);
        Assert.Equal(ChatCommands.ChatSlot, routing.ServiceSlot);
        Assert.Equal(16, reader.RemainingBits + reader.Position);
    }

    [Fact]
    public void Ping_WritesThePresenceBitThenBothTimestampHalves()
    {
        var reader = new BitReader(ConnectionCommands.Ping(0x0000_0102_0304_0506));
        var routing = RoutingHeader.Decode(reader);

        Assert.Equal((ConnectionCommands.ConnectionSlot, ConnectionCommands.PingCommand), (routing.ServiceSlot!.Value, routing.CommandId));
        Assert.Equal(1UL, reader.Read(1));
        Assert.Equal(0x0000_0102UL, reader.Read(32));
        Assert.Equal(0x0304_0506UL, reader.Read(32));
        Assert.Equal(10 * 8, reader.Position + reader.RemainingBits);
    }

    [Fact]
    public void ServerPing_DecodesAndThePongEchoesTheSameBytes()
    {
        var timestamp = new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 };
        var ping = Record(ConnectionCommands.ConnectionSlot, ConnectionCommands.PingCommand, w =>
        {
            w.Write(1, 1);
            w.WriteBytes(timestamp, aligned: true);
        });

        var decoded = Assert.IsType<NativeChatRecord.Ping>(DecodeOne(ping));
        Assert.Equal(timestamp, decoded.Timestamp);

        var pong = ConnectionCommands.Pong(decoded.Timestamp);
        var echoed = Assert.IsType<NativeChatRecord.Pong>(DecodeOne(pong));
        Assert.Equal(timestamp, echoed.Timestamp);
        Assert.Equal(ping.AsSpan(2).ToArray(), pong.AsSpan(2).ToArray());
    }

    [Fact]
    public void ServerPing_WithoutATimestamp_IsTwoBytesAndPongsTheSame()
    {
        var ping = Record(ConnectionCommands.ConnectionSlot, ConnectionCommands.PingCommand, w => w.Write(0, 1));

        var decoded = Assert.IsType<NativeChatRecord.Ping>(DecodeOne(ping));

        Assert.Null(decoded.Timestamp);
        Assert.Equal(2, ping.Length);
        Assert.Equal(2, ConnectionCommands.Pong(null).Length);
    }

    [Fact]
    public void GameSiteInfo_ReadsEverySiteToTheExactEnd()
    {
        var record = Record(1, 14, w =>
        {
            w.Write(0x7F00_0001, 32);
            w.Write(1119, 16);
            w.Align();
            w.Write(2, 7);
            WriteBlob(w, 6, "US10-S2");
            w.Write(1, 1);
            w.Align();
            w.Write(0x0A00_0001, 32);
            w.Write(1119, 16);
            WriteBlob(w, 6, "SG1");
            w.Write(0, 1);
        });

        AssertDecodesThenMarker<NativeChatRecord.Sc2ServerCatalog>(record);
    }

    [Fact]
    public void ToonWelcome_ReadsToTheExactEndAndCountsUnlocks()
    {
        var record = Record(ChatCommands.ToonSlot, 10, w =>
        {
            w.Write(0, 32);
            w.Write(1, 4); // one cache handle
            w.WriteBytes(new byte[40], aligned: true);
            w.Write(0, 32);
            w.Write(0, 1);
            w.Write(0, 32);
            w.Write(0, 31);
            w.Write(0, 32);
            WriteBlob(w, 6, "x");
            w.Write(0, 64);
            w.Write(0, 64);
            w.Write(1, 3); // one realm map
            w.Write(0, 32);
            w.Write(0, 32);
            w.Write(0, 1);
            w.Write(0, 8);
            w.Write(2, 5);
            w.Write(0, 32);
            w.Write(0, 32);
            w.Write(3, 8); // three unlocks
            for (var i = 0; i < 3; i++)
            {
                w.Write(0, 4);
                w.WriteBytes(new byte[40], aligned: true);
            }

            w.Write(0, 3);
            w.Write(0, 16);
            WriteBlob(w, 13, "a");
            WriteBlob(w, 13, "bc");
            w.Write(0, 32);
        });

        var welcome = AssertDecodesThenMarker<NativeChatRecord.Sc2ToonWelcomeUnlocks>(record);
        Assert.Equal(3, welcome.UnlockCount);
    }

    [Fact]
    public void MessageFrame_ReadsEveryHeaderKind()
    {
        var record = Record(1, 13, w =>
        {
            w.Write(3, 14);
            w.WriteBytes([1, 2, 3], aligned: true);
            w.Write(0, 8);
            w.Write(9, 6);
            w.Write(0, 8); w.Write(0, 64);
            w.Write(1, 8); w.Write(0, 32); w.Write(0, 32); w.Write(0, 6); w.Write(1, 1); w.Write(0, 64);
            w.Write(2, 8); w.Write(0, 7); w.Write(3, 3); w.Write(2, 11); w.Write(0, 64); w.Write(0, 64); w.Write(0, 16);
            w.Write(3, 8); w.Write(0, 33);
            w.Write(5, 8); w.Write(0, 64); w.Write(0, 32);
            w.Write(6, 8); w.Write(0, 16); WriteBlob(w, 14, "oops");
            w.Write(7, 8); w.Write(1, 1);
            w.Write(8, 8); w.Write(0, 32);
            w.Write(9, 8); w.Write(0, 17);
        });

        AssertDecodesThenMarker<NativeChatRecord.Sc2Consumed>(record);
    }

    [Fact]
    public void PresenceUpdate_ReadsToTheExactEnd()
    {
        var record = Record(4, 0, w =>
        {
            w.Write(0, 19);
            w.Write(1, 1);
            w.Write(0, 64);
            WriteBlob(w, 11, "status");
            w.Write(0, 11);
            w.Write(2, 4); w.Write(0, 64);
            w.Write(0, 4);
            w.Write(1, 4); w.Write(0, 16);
            w.Write(1, 1); w.Write(0, 33);
            w.Write(0, 8);
        });

        var update = AssertDecodesThenMarker<NativeChatRecord.PresenceUpdate>(record);
        Assert.Equal("status", System.Text.Encoding.UTF8.GetString(update.Value.FieldData));
    }

    [Fact]
    public void PartyRecord_SkipsToItsFixedEighteenBytes()
    {
        var record = Record(12, 0, w => w.WriteBytes(new byte[18 - 2], aligned: true));

        Assert.Equal(18, record.Length);
        AssertDecodesThenMarker<NativeChatRecord.Sc2Consumed>(record);
    }

    [Fact]
    public void AccountBlocks_ReadFullNamesAndStayInStep()
    {
        var record = Record(3, 31, w =>
        {
            w.Write(1, 1); w.Write(1, 1);
            w.Write(1, 7);
            w.Write(0, 9);
            w.Write(0, 32);
            w.Write(1, 1); WriteBlob(w, 7, "Jimmy");
            w.Write(0, 20);
            w.Write(1, 1); WriteBlob(w, 8, "Jim"); WriteBlob(w, 8, "Raynor");
            w.Write(0, 32);
        });

        AssertDecodesThenMarker<NativeChatRecord.Sc2Consumed>(record);
    }

    private static byte[] Record(byte slot, byte command, Action<BitWriter> payload)
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, command, slot);
        payload(writer);
        writer.Align();
        return writer.ToBytes();
    }

    private static void WriteBlob(BitWriter writer, int lengthBits, string text)
    {
        var bytes = System.Text.Encoding.UTF8.GetBytes(text);
        writer.Write((ulong)bytes.Length, lengthBits);
        writer.WriteBytes(bytes, aligned: true);
    }

    private static NativeChatRecord DecodeOne(byte[] record)
    {
        var reader = new BitReader(record);
        var routing = RoutingHeader.Decode(reader);
        var decoded = NativeRecordDispatcher.Decode(routing.CommandId, routing.ServiceSlot, reader);
        reader.Align();
        Assert.Equal(record.Length * 8, reader.Position);
        return decoded;
    }

    /// <summary>Decodes <paramref name="record"/> through a RecordStream with a chat message right behind it, which only decodes if the first record ended on exactly the right byte.</summary>
    private static T AssertDecodesThenMarker<T>(byte[] record)
        where T : NativeChatRecord
    {
        // A server-sent chat message (sender handle first), not the client's own
        // Chat 11 layout, which has no handle.
        var marker = Record(ChatCommands.ChatSlot, 11, w =>
        {
            w.Write(42, 32);
            WriteBlob(w, 10, "marker");
            w.Write(2, 3);
        });
        using var stream = new RecordStream(new MemoryStream([.. record, .. marker]));
        stream.FillAsync().GetAwaiter().GetResult();

        Assert.True(stream.TryDecodeRecord(NativeRecordDispatcher.Decode, out var first));
        Assert.True(stream.TryDecodeRecord(NativeRecordDispatcher.Decode, out var second));
        var message = Assert.IsType<NativeChatRecord.Message>(second);
        Assert.Equal("marker", message.Value.Body);
        Assert.Equal((byte)2, message.Value.ChannelIndex);
        return Assert.IsType<T>(first);
    }
}
