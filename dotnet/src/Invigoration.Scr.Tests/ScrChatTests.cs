using Invigoration.Sc2.Protobuf;
using Invigoration.Scr.Aurora;
using Invigoration.Scr.Classic;
using Invigoration.Scr.LegacyChat;

namespace Invigoration.Scr.Tests;

/// <summary>
/// The byte example is the published one on docs.bnet.cc's "StarCraft:
/// Remastered chat" ("hello" to channel 9, token 1, seed 0x12345678). There are
/// no captures of server calls yet, so those are built here from the
/// documented field numbers.
/// </summary>
public class ScrChatTests
{
    private const uint ExampleSeed = 0x12345678;
    private const string ExamplePlain = "001a08f8b484a70f10adf6c7c205180120f1c988b4092809300048000809120568656c6c6f";
    private const string ExampleWire = "f043aadb24517cee8fdad2652ea675d30ebd1f1364feef00a9ffa701f7ae11a1352364164e";

    [Fact]
    public void SendMessage_BuildsThePublishedPlainFrame()
    {
        var (method, body) = LegacyChatRequests.SendMessage(9, "hello");
        var header = new ClassicHeader(LegacyChatService.Hash, method, 1, ClassicHeader.RequestRouting, 0, 0, false);

        Assert.Equal(ExamplePlain, Hex(ClassicFrame.Encode(header, body)));
    }

    [Fact]
    public void Envelope_ScramblesAndUnscramblesThePublishedExample()
    {
        Assert.Equal(ExampleWire, Hex(ClassicEnvelope.Scramble(Convert.FromHexString(ExamplePlain), ExampleSeed)));
        Assert.Equal(ExamplePlain, Hex(ClassicEnvelope.Unscramble(Convert.FromHexString(ExampleWire), ExampleSeed)));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(7)]
    [InlineData(33)]
    public void Envelope_RoundTripsAnyLength(int length)
    {
        var plain = Enumerable.Range(0, length).Select(i => (byte)(i * 37)).ToArray();

        Assert.Equal(plain, ClassicEnvelope.Unscramble(ClassicEnvelope.Scramble(plain, 0xDEADBEEF), 0xDEADBEEF));
    }

    [Fact]
    public void Session_FirstMessageIsTheWholePublishedWireExample()
    {
        var session = new ScrChatSession(ExampleSeed, token: 0, channelId: 9);

        Assert.Equal(ExampleWire, Hex(session.SendMessage("hello")));
    }

    [Fact]
    public void Whisper_KeepsTheWholeTextInOneArgument()
    {
        var (method, body) = LegacyChatRequests.Whisper(9, "Raynor", "meet at the bar");

        Assert.Equal(LegacyChatService.CommandMethod, method);
        var fields = Fields(body);
        Assert.Equal([(1, "9"), (2, "whisper"), (3, "Raynor"), (3, "meet at the bar")], fields);
    }

    [Fact]
    public void JoinByName_SendsTheNameAsOneArgument()
    {
        var (method, body) = LegacyChatRequests.JoinByName(9, "Clan Raiders");

        Assert.Equal(LegacyChatService.CommandMethod, method);
        Assert.Equal([(1, "9"), (2, "channel"), (3, "Clan Raiders")], Fields(body));
    }

    [Fact]
    public void Frames_SeveralRpcsInOneMessageAllDecode()
    {
        var first = ClassicFrame.Encode(Call(LegacyChatService.Hash, 0x850B6EE3, 5), [1, 2]);
        var second = ClassicFrame.Encode(Call(LegacyChatService.Hash, 0x632D6CFD, 6), [3]);

        var rpcs = ClassicFrame.DecodeAll([.. first, .. second]);

        Assert.Equal([5u, 6u], rpcs.Select(r => r.Header.Token));
        Assert.Equal([3], rpcs[1].Body);
    }

    [Fact]
    public void Frames_MessageCutShortIsRejected()
    {
        var frame = ClassicFrame.Encode(Call(LegacyChatService.Hash, 0x850B6EE3, 5), [1, 2, 3]);

        Assert.Throws<InvalidOperationException>(() => ClassicFrame.DecodeAll(frame.AsSpan(0, frame.Length - 1)));
    }

    [Fact]
    public void Session_AnswersEveryCallAndOnlyMovesOnAConfirmedChannel()
    {
        var session = new ScrChatSession(ExampleSeed, token: 10, channelId: 9);
        session.JoinByName("Clan Raiders");
        Assert.Equal(9ul, session.ChannelId);

        var channel = new ProtoWriter();
        channel.WriteUInt64(1, 42);
        channel.WriteString(2, "clan raiders");
        channel.WriteBytesField(3, Member("Raynor", ("race", "Terran")));
        channel.WriteString(5, "Clan Raiders");
        var current = new ProtoWriter();
        current.WriteBytesField(2, channel.ToArray());
        var call = ClassicFrame.Encode(Call(LegacyChatService.Hash, LegacyChatService.CurrentChannelCallback, 77), current.ToArray());

        var (events, replies) = session.Receive(ClassicEnvelope.Scramble(call, ExampleSeed));

        var entered = Assert.IsType<ScrChatEvent.ChannelEntered>(Assert.Single(events));
        Assert.Equal("Clan Raiders", entered.Channel.DisplayName);
        Assert.Equal("Terran", Assert.Single(entered.Channel.Members).Attributes["race"]);
        Assert.Equal(42ul, session.ChannelId);

        var reply = Assert.Single(ClassicFrame.DecodeAll(ClassicEnvelope.Unscramble(Assert.Single(replies), ExampleSeed)));
        Assert.True(reply.Header.IsResponse);
        Assert.Equal((LegacyChatService.Hash, LegacyChatService.CurrentChannelCallback, 77u, ClassicHeader.RequestRouting), (reply.Header.Service, reply.Header.Method, reply.Header.Token, reply.Header.Routing!.Value));
        Assert.Empty(reply.Body);
    }

    [Fact]
    public void Session_EchoesConnectionServiceBodies()
    {
        var session = new ScrChatSession(ExampleSeed, token: 0, channelId: 9);
        var call = ClassicFrame.Encode(Call(LegacyChatService.ConnectionHash, LegacyChatService.ConnectionEchoMethod, 3), [9, 8, 7]);

        var (events, replies) = session.Receive(ClassicEnvelope.Scramble(call, ExampleSeed));

        Assert.Empty(events);
        var reply = Assert.Single(ClassicFrame.DecodeAll(ClassicEnvelope.Unscramble(Assert.Single(replies), ExampleSeed)));
        Assert.Equal([9, 8, 7], reply.Body);
    }

    [Fact]
    public void Session_ReportsAFailedRequestByItsMethod()
    {
        var session = new ScrChatSession(ExampleSeed, token: 0, channelId: 9);
        session.SendMessage("hi");
        var failure = ClassicFrame.Encode(
            new ClassicHeader(LegacyChatService.Hash, LegacyChatService.SendMessageMethod, 1, ClassicHeader.RequestRouting, 0, 13, true), []);

        var (events, replies) = session.Receive(ClassicEnvelope.Scramble(failure, ExampleSeed));

        Assert.Empty(replies);
        Assert.Equal(new ScrChatEvent.RequestFailed(LegacyChatService.SendMessageMethod, 1, 13), Assert.Single(events));
    }

    [Fact]
    public void Message_FirstTextIsTheSenderAndLastIsTheText()
    {
        var inner = new ProtoWriter();
        inner.WriteString(1, "Raynor");
        inner.WriteUInt64(2, 5);
        inner.WriteString(3, "hello there");
        var outer = new ProtoWriter();
        outer.WriteUInt64(1, 9);
        outer.WriteBytesField(2, inner.ToArray());

        var message = LegacyChatCallbacks.DecodeMessage(ScrMessageKind.Channel, outer.ToArray());

        Assert.Equal(new ScrMessage(ScrMessageKind.Channel, "Raynor", "hello there"), message);
    }

    [Fact]
    public void ChannelListChanges_MarkTypeOneAsARemoval()
    {
        var removed = new ProtoWriter();
        removed.WriteUInt64(1, 1);
        removed.WriteBytesField(2, Channel(7, "Old"));
        var added = new ProtoWriter();
        added.WriteUInt64(1, 0);
        added.WriteBytesField(2, Channel(8, "New"));
        var body = new ProtoWriter();
        body.WriteBytesField(1, removed.ToArray());
        body.WriteBytesField(1, added.ToArray());

        var changes = LegacyChatCallbacks.DecodeChannelListChanges(body.ToArray());

        Assert.Equal([(true, 7ul), (false, 8ul)], changes.Select(c => (c.IsRemoval, c.Channel.Id)));
    }

    [Fact]
    public void Aurora_EchoIsAnsweredWithTheSameTokenAndBody()
    {
        var request = AuroraKeepalive.Handle("""[{"service_hash":1698982289,"method_id":3,"token":12},{"payload":"abc"}]""", out var reply);

        Assert.Equal(AuroraKeepalive.Request.Echo, request);
        Assert.Equal("""[{"service_id":254,"token":12,"is_response":true,"status":0},{"payload":"abc"}]""", reply);
    }

    [Fact]
    public void Aurora_DisconnectAndOtherCallsAreRecognised()
    {
        Assert.Equal(AuroraKeepalive.Request.Disconnect, AuroraKeepalive.Handle("""[{"service_hash":1698982289,"method_id":4,"token":1},{}]""", out _));
        Assert.Equal(AuroraKeepalive.Request.None, AuroraKeepalive.Handle("""[{"service_hash":1,"method_id":3,"token":1},{}]""", out var reply));
        Assert.Null(reply);
    }

    private static ClassicHeader Call(uint service, uint method, uint token) =>
        new(service, method, token, ClassicHeader.RequestRouting, 0, 0, false);

    private static byte[] Channel(ulong id, string name)
    {
        var w = new ProtoWriter();
        w.WriteUInt64(1, id);
        w.WriteString(5, name);
        return w.ToArray();
    }

    private static byte[] Member(string name, (string Name, string Value) attribute)
    {
        var a = new ProtoWriter();
        a.WriteString(1, attribute.Name);
        a.WriteString(2, attribute.Value);
        var w = new ProtoWriter();
        w.WriteString(1, name);
        w.WriteUInt64(2, 0);
        w.WriteBytesField(3, a.ToArray());
        return w.ToArray();
    }

    private static List<(int Field, string Value)> Fields(byte[] body)
    {
        var result = new List<(int, string)>();
        var r = new ProtoReader(body);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            result.Add((field, type == WireType.Varint ? r.ReadVarint().ToString() : r.ReadString()));
        }

        return result;
    }

    private static string Hex(byte[] bytes) => Convert.ToHexString(bytes).ToLowerInvariant();
}
