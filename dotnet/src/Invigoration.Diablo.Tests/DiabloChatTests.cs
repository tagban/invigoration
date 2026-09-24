using Invigoration.Sc2.Front;
using Invigoration.Sc2.Protobuf;

namespace Invigoration.Diablo.Tests;

/// <summary>
/// The two send-message frames are the published examples on docs.bnet.cc's
/// "Diablo II: Resurrected and Diablo IV chat" (game account 456, channel 99 on
/// host 1 / epoch 2, region 1, token 1, "hello"). There are no captures of
/// server calls yet, so those are built here from the documented field numbers.
/// </summary>
public class DiabloChatTests
{
    private static readonly GameAccount D2rAccount = new(456, 0x4F5349, 1);
    private static readonly GameAccount D4Account = new(456, 0x46656E, 1);
    private static readonly DiabloChannelId ExampleChannel = new(1, 2, 99, 1);

    [Fact]
    public void D2r_SendMessage_IsThePublishedFrame()
    {
        var session = new DiabloChatSession(DiabloGame.D2R, D2rAccount, ExampleChannel, token: 0);

        Assert.Equal(
            "000d08001017180128265dd1398d790a0c0dc80100001549534f001801120d1204080110021d6300000020011a07220568656c6c6f",
            Hex(session.SendMessage("hello")));
    }

    [Fact]
    public void D4_SendMessage_IsThePublishedFrame()
    {
        var session = new DiabloChatSession(DiabloGame.D4, D4Account, ExampleChannel, token: 0);

        Assert.Equal(
            "000d08001017180128245dd1398d79120d1204080110021d6300000020011a07220568656c6c6f220a08c80310eeca99021801",
            Hex(session.SendMessage("hello")));
    }

    [Fact]
    public void ServiceHashes_MatchTheDocumentedValues()
    {
        Assert.Equal(0x798D39D1u, DiabloServices.ChannelService);
        Assert.Equal(0x1AE52686u, DiabloServices.ChannelListener);
        Assert.Equal(0x018007BEu, DiabloServices.MembershipListener);
        Assert.Equal(0xFEE1AA14u, DiabloServices.WhisperService);
        Assert.Equal(0x62615E21u, DiabloServices.WhisperListener);
        Assert.Equal(0x8B2A82A2u, DiabloServices.AccountService);
        Assert.Equal(0x65446991u, DiabloServices.ConnectionService);
    }

    [Fact]
    public void SendMessage_RejectsTextOver255Bytes()
    {
        var session = new DiabloChatSession(DiabloGame.D2R, D2rAccount, ExampleChannel, token: 0);

        Assert.Throws<ArgumentException>(() => session.SendMessage(new string('é', 128)));
    }

    [Theory]
    [InlineData(DiabloGame.D2R)]
    [InlineData(DiabloGame.D4)]
    public void Message_IsNamedFromTheRosterAndOtherChannelsAreIgnored(DiabloGame game)
    {
        var account = game == DiabloGame.D2R ? D2rAccount : D4Account;
        var requests = new DiabloChatRequests(game, account);
        var raynor = new GameAccount(456, account.TitleId, 1);
        var session = new DiabloChatSession(game, account, ExampleChannel, token: 0, [new DiabloMember(raynor, "Raynor", 7)]);

        var (events, _) = session.Receive(Call(DiabloServices.ChannelListener, 10, ChatBody(requests, ExampleChannel, "hi all")));
        var (elsewhere, _) = session.Receive(Call(DiabloServices.ChannelListener, 10, ChatBody(requests, ExampleChannel with { Id = 5 }, "not here")));

        var message = Assert.IsType<DiabloChatEvent.Message>(Assert.Single(events));
        Assert.Equal(("Raynor", "hi all"), (message.SenderName, message.Value.Text));
        Assert.Empty(elsewhere);
    }

    [Fact]
    public void Roster_TracksJoinsAndLeavesUsingD4sHandleField()
    {
        var session = new DiabloChatSession(DiabloGame.D4, D4Account, ExampleChannel, token: 0);
        var requests = new DiabloChatRequests(DiabloGame.D4, new GameAccount(789, D4Account.TitleId, 1));
        var member = new ProtoWriter();
        member.WriteString(2, "Nova");
        member.WriteUInt64(6, 42);
        member.WriteBytesField(7, requests.Handle());

        var (joined, _) = session.Receive(Call(DiabloServices.ChannelListener, 3, InChannel(ExampleChannel, 4, member.ToArray())));
        var (left, _) = session.Receive(Call(DiabloServices.ChannelListener, 4, InChannel(ExampleChannel, 4, requests.Handle())));

        Assert.Equal("Nova", Assert.IsType<DiabloChatEvent.MemberJoined>(Assert.Single(joined)).Member.Name);
        Assert.Equal("Nova", Assert.IsType<DiabloChatEvent.MemberLeft>(Assert.Single(left)).Member.Name);
        Assert.Empty(session.Members);
    }

    [Fact]
    public void JoinChannel_SubscribesOnlyOnceFindReplyAndDescriptionHaveBothArrived()
    {
        var session = new DiabloChatSession(DiabloGame.D2R, D2rAccount, ExampleChannel, token: 0);
        var type = new UniqueChannelType(0x4F5349, "public_default");
        var target = new DiabloChannelId(1, 2, 100, 1);
        session.JoinChannel(type, "trade-1");

        var description = new ProtoWriter();
        description.WriteBytesField(1, DiabloChatRequests.ChannelId(target));
        description.WriteString(3, "Trade 1");
        var identity = new ProtoWriter();
        identity.WriteString(1, "trade-1");
        description.WriteBytesField(110, identity.ToArray());
        var membership = new ProtoWriter();
        membership.WriteBytesField(3, description.ToArray());

        var (_, afterDescription) = session.Receive(Call(DiabloServices.MembershipListener, 1, membership.ToArray()));
        Assert.Empty(afterDescription);

        var (_, afterFind) = session.Receive(Reply(token: 1, []));
        var subscribe = FrontFrame.Decode(Assert.Single(afterFind));
        Assert.Equal((DiabloServices.ChannelService, 10u, 2u), (subscribe.Header.ServiceHash!.Value, subscribe.Header.MethodId!.Value, subscribe.Header.Token));
        Assert.Equal(ExampleChannel, session.Channel);

        var (joined, _) = session.Receive(Reply(token: 2, []));
        Assert.Equal("Trade 1", Assert.IsType<DiabloChatEvent.ChannelJoined>(Assert.Single(joined)).Channel.Name);
        Assert.Equal(target, session.Channel);
    }

    [Fact]
    public void Whisper_LooksUpTheBattleTagThenSends()
    {
        var session = new DiabloChatSession(DiabloGame.D4, D4Account, ExampleChannel, token: 0);
        session.Whisper("Raynor#1234", "hello");
        var resolved = new ProtoWriter();
        var nested = new ProtoWriter();
        nested.WriteUInt64(1, 555);
        resolved.WriteBytesField(12, nested.ToArray());

        var (_, outgoing) = session.Receive(Reply(token: 1, resolved.ToArray()));

        var whisper = FrontFrame.Decode(Assert.Single(outgoing));
        Assert.Equal((DiabloServices.WhisperService, 4u), (whisper.Header.ServiceHash!.Value, whisper.Header.MethodId!.Value));
        Assert.Equal(DiabloChatRequests.Whisper(555, "hello"), whisper.Body);
    }

    [Fact]
    public void ResolvedAccountId_AlsoReadsTheOlderUnnestedForm()
    {
        var older = new ProtoWriter();
        older.WriteUInt64(12, 777);

        Assert.Equal(777ul, DiabloChatDecoders.DecodeResolvedAccountId(older.ToArray()));
    }

    [Fact]
    public void IncomingWhisper_ReadsTextSenderAndBattleTag()
    {
        var session = new DiabloChatSession(DiabloGame.D2R, D2rAccount, ExampleChannel, token: 0);
        var whisper = new ProtoWriter();
        whisper.WriteUInt64(2, 555);
        whisper.WriteString(5, "psst");
        var body = new ProtoWriter();
        body.WriteBytesField(2, whisper.ToArray());
        body.WriteString(3, "Raynor#1234");

        var (events, _) = session.Receive(Call(DiabloServices.WhisperListener, 1, body.ToArray()));

        Assert.Equal(new DiabloWhisper(555, "Raynor#1234", "psst"), Assert.IsType<DiabloChatEvent.Whisper>(Assert.Single(events)).Value);
    }

    [Fact]
    public void ConnectionEcho_IsAnsweredWithTheDocumentedFields()
    {
        var session = new DiabloChatSession(DiabloGame.D2R, D2rAccount, ExampleChannel, token: 0);
        var echo = new ProtoWriter();
        echo.WriteFixed64(1, 0x1122334455667788);
        echo.WriteBytesField(3, [1, 2, 3]);

        var (_, outgoing) = session.Receive(FrontFrame.Encode(
            new Header { ServiceId = 0, MethodId = 3, Token = 9, ServiceHash = DiabloServices.ConnectionService }, echo.ToArray()));

        var reply = FrontFrame.Decode(Assert.Single(outgoing));
        Assert.Equal((254u, 9u, true), (reply.Header.ServiceId, reply.Header.Token, reply.Header.IsResponse!.Value));
        var fields = ProtoFields.Parse(reply.Body);
        Assert.Equal(0x1122334455667788ul, fields.Number(1));
        Assert.Equal([1, 2, 3], fields.Bytes(2));
    }

    [Fact]
    public void KeepAlive_IsConnectionMethodFiveWithAnEmptyBody()
    {
        var session = new DiabloChatSession(DiabloGame.D2R, D2rAccount, ExampleChannel, token: 0);

        var (header, body) = FrontFrame.Decode(session.KeepAlive());

        Assert.Equal((DiabloServices.ConnectionService, 5u, 1u), (header.ServiceHash!.Value, header.MethodId!.Value, header.Token));
        Assert.Empty(body);
    }

    private static byte[] ChatBody(DiabloChatRequests sender, DiabloChannelId channel, string text)
    {
        var message = new ProtoWriter();
        message.WriteBytesField(1, sender.Handle());
        message.WriteString(3, text);
        return InChannel(channel, 4, message.ToArray());
    }

    private static byte[] InChannel(DiabloChannelId channel, int field, byte[] payload)
    {
        var w = new ProtoWriter();
        w.WriteBytesField(3, DiabloChatRequests.ChannelId(channel));
        w.WriteBytesField(field, payload);
        return w.ToArray();
    }

    private static byte[] Call(uint service, uint method, byte[] body) =>
        FrontFrame.Encode(new Header { ServiceId = 0, MethodId = method, Token = 1000, ServiceHash = service }, body);

    private static byte[] Reply(uint token, byte[] body) =>
        FrontFrame.Encode(new Header { ServiceId = 254, Token = token, Status = 0, IsResponse = true }, body);

    private static string Hex(byte[] bytes) => Convert.ToHexString(bytes).ToLowerInvariant();
}
