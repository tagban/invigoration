using System.Text;
using Invigoration.Sc2.Native;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Tests;

public class RecordStreamTests
{
    /// <summary>Builds server->client bytes for MessageRecv (command 11) — a different, larger payload shape than the client->server ChatCommands.ChatMessage builder for the same command/slot, since route dispatch is keyed by direction too.</summary>
    private static byte[] BuildInboundMessageRecvRecord(uint memberHandle, string body, byte channelIndex)
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, commandId: 11, serviceSlot: ChatCommands.ChatSlot);
        writer.Write(memberHandle, 32);
        var bodyBytes = Encoding.UTF8.GetBytes(body);
        writer.Write((ulong)bodyBytes.Length, 10);
        writer.WriteBytes(bodyBytes, aligned: true);
        writer.Write(channelIndex, 3);
        writer.Align();
        return writer.ToBytes();
    }

    /// <summary>Builds server->client bytes for WhisperRecv (command 19) — the peer-identifying shape, unlike the client->server ChatCommands.ChatWhisper builder for the same command/slot.</summary>
    private static byte[] BuildInboundWhisperRecvRecord(byte region, uint programId, uint realm, string peerName, string body)
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, commandId: 19, serviceSlot: ChatCommands.ChatSlot);
        writer.Write(region, 8);
        writer.Write(programId, 32);
        writer.Write(realm, 32);
        var nameBytes = Encoding.UTF8.GetBytes(peerName);
        writer.Write((ulong)(nameBytes.Length - 2), 7);
        writer.WriteBytes(nameBytes, aligned: true);
        var bodyBytes = Encoding.UTF8.GetBytes(body);
        writer.Write((ulong)bodyBytes.Length, 10);
        writer.WriteBytes(bodyBytes, aligned: true);
        writer.Align();
        return writer.ToBytes();
    }

    [Fact]
    public async Task SendAsync_WithoutEncryption_WritesRawBytes()
    {
        using var ms = new MemoryStream();
        var stream = new RecordStream(ms);
        var record = ChatCommands.ChatLeave(0);

        await stream.SendAsync(record);

        Assert.Equal(record, ms.ToArray());
    }

    [Fact]
    public async Task SendAsync_WithEncryption_EncryptsWithOutboundCipher()
    {
        using var ms = new MemoryStream();
        var stream = new RecordStream(ms);
        var key = new byte[] { 1, 2, 3, 4 };
        stream.EnableEncryption(new Rc4State(key), new Rc4State(key));
        var record = ChatCommands.ChatLeave(0);

        await stream.SendAsync(record);

        var expected = new Rc4State(key).Apply(record);
        Assert.Equal(expected, ms.ToArray());
        Assert.NotEqual(record, ms.ToArray());
    }

    [Fact]
    public async Task FillAsync_ThenTryDecodeRecord_DecodesBufferedRecord()
    {
        var record = ChatCommands.ChatLeave(3);
        using var ms = new MemoryStream(record);
        var stream = new RecordStream(ms);

        var filled = await stream.FillAsync();
        Assert.True(filled);

        var decoded = stream.TryDecodeRecord(
            (command, slot, reader) => (command, slot, channelIndex: reader.Read(3)),
            out var result);

        Assert.True(decoded);
        Assert.Equal(3u, result.channelIndex);
    }

    [Fact]
    public async Task TryDecodeRecord_WithPartialData_ReturnsFalseUntilFullyBuffered()
    {
        var record = BuildInboundMessageRecvRecord(memberHandle: 99, body: "hello there", channelIndex: 0);
        using var ms = new SegmentedReadStream(record, firstChunkSize: 2);
        var stream = new RecordStream(ms);

        await stream.FillAsync();
        var decodedEarly = stream.TryDecodeRecord(
            (command, slot, reader) => ChatRecordDecoder.DecodeChatMessage(reader),
            out _);
        Assert.False(decodedEarly);

        await stream.FillAsync();
        var decoded = stream.TryDecodeRecord(
            (command, slot, reader) => ChatRecordDecoder.DecodeChatMessage(reader),
            out var message);

        Assert.True(decoded);
        Assert.Equal("hello there", message!.Body);
    }

    [Fact]
    public async Task RoundTrip_ThroughRc4Encryption_DecodesCorrectly()
    {
        var key = new byte[] { 9, 8, 7, 6, 5 };
        var record = BuildInboundWhisperRecvRecord(region: 1, programId: FourCc.Encode("S2"), realm: 1, peerName: "Tagban#1234", body: "hi");
        var encrypted = new Rc4State(key).Apply(record);

        using var ms = new MemoryStream(encrypted);
        var stream = new RecordStream(ms);
        stream.EnableEncryption(new Rc4State(key), new Rc4State(key));

        await stream.FillAsync();
        var decoded = stream.TryDecodeRecord(
            (command, slot, reader) => ChatRecordDecoder.DecodeChatWhisper(reader),
            out var whisper);

        Assert.True(decoded);
        Assert.Equal("hi", whisper!.Body);
    }

    /// <summary>A stream that hands back its bytes in two reads, simulating TCP segmentation of a single record.</summary>
    private sealed class SegmentedReadStream(byte[] data, int firstChunkSize) : Stream
    {
        private int _position;

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => data.Length;
        public override long Position { get => _position; set => throw new NotSupportedException(); }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (_position >= data.Length)
            {
                return 0;
            }

            var remaining = data.Length - _position;
            var take = _position == 0 ? Math.Min(firstChunkSize, remaining) : remaining;
            Array.Copy(data, _position, buffer, offset, take);
            _position += take;
            return take;
        }

        public override void Flush()
        {
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
