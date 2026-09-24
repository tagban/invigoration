using Invigoration.Sc2.Protobuf;
using Invigoration.Sc2.Wire;

namespace Invigoration.Diablo;

/// <summary>Which game's variant of the chat protocol to speak.</summary>
public enum DiabloGame
{
    /// <summary>Diablo II: Resurrected: fixed-width handles, Handle first in bodies.</summary>
    D2R,

    /// <summary>Diablo IV: varint handles, Handle last in bodies.</summary>
    D4,
}

/// <summary>The signed-in game account: its ID, title ID and region.</summary>
public sealed record GameAccount(ulong Id, uint TitleId, uint Region);

/// <summary>A channel. All four parts are needed to address it; the numeric ID alone isn't enough.</summary>
public sealed record DiabloChannelId(ulong HostLabel, ulong HostEpoch, uint Id, uint Region);

/// <summary>A public channel type: the game's program code plus a type name such as "public_default".</summary>
public sealed record UniqueChannelType(uint Program, string ChannelType);

/// <summary>Front service names and hashes used for D2R and D4 chat.</summary>
public static class DiabloServices
{
    public const string ChannelServiceName = "bnet.protocol.channel.v2.ChannelService";
    public const string ChannelListenerName = "bnet.protocol.channel.v2.ChannelListener";
    public const string MembershipListenerName = "bnet.protocol.channel.v2.membership.ChannelMembershipListener";
    public const string WhisperServiceName = "bnet.protocol.whisper.v2.client.WhisperService";
    public const string WhisperListenerName = "bnet.protocol.whisper.v2.client.WhisperListener";
    public const string AccountServiceName = "bnet.protocol.account.v1.AccountService";
    public const string ConnectionServiceName = "bnet.protocol.connection.ConnectionService";

    public static readonly uint ChannelService = ServiceHash.Compute(ChannelServiceName);
    public static readonly uint ChannelListener = ServiceHash.Compute(ChannelListenerName);
    public static readonly uint MembershipListener = ServiceHash.Compute(MembershipListenerName);
    public static readonly uint WhisperService = ServiceHash.Compute(WhisperServiceName);
    public static readonly uint WhisperListener = ServiceHash.Compute(WhisperListenerName);
    public static readonly uint AccountService = ServiceHash.Compute(AccountServiceName);
    public static readonly uint ConnectionService = ServiceHash.Compute(ConnectionServiceName);

    public const uint SendMessageMethod = 23;
    public const uint ListChannelTypesMethod = 5;
    public const uint FindChannelMethod = 6;
    public const uint SubscribeMethod = 10;
    public const uint UnsubscribeMethod = 11;
    public const uint LeaveMethod = 31;
    public const uint ResolveBattleTagMethod = 13;
    public const uint SendWhisperMethod = 4;
    public const uint KeepAliveMethod = 5;
    public const uint EchoMethod = 3;
    public const uint DisconnectRequestMethod = 4;

    public const uint MessageCallback = 10;
    public const uint MemberAddedCallback = 3;
    public const uint MemberRemovedCallback = 4;
    public const uint WhisperCallback = 1;
    public const uint ChannelDescriptionCallback = 1;
}

/// <summary>
/// Request bodies for D2R and D4 chat over Front RPC, after sign-in. D2R and D4
/// carry the same fields but encode the Handle differently and put it in a
/// different place. See docs.bnet.cc's "Diablo II: Resurrected and Diablo IV chat".
/// </summary>
public sealed class DiabloChatRequests(DiabloGame game, GameAccount account)
{
    /// <summary>Most chat text allowed, in UTF-8 bytes.</summary>
    public const int MaxMessageBytes = 255;

    private const uint EnUsLocale = 0x656E5553; // 'enUS'

    public DiabloGame Game { get; } = game;

    /// <summary>The game's program code as a big-endian ASCII number: 'OSI' for D2R, 'Fen' for D4.</summary>
    public uint Program => Game == DiabloGame.D2R ? 0x4F5349u : 0x46656Eu;

    /// <summary>The public channel type to list: "public_default" for D2R, "public_test" for D4.</summary>
    public string PublicChannelType => Game == DiabloGame.D2R ? "public_default" : "public_test";

    public byte[] Handle()
    {
        var w = new ProtoWriter();
        if (Game == DiabloGame.D2R)
        {
            w.WriteFixed32(1, (uint)(account.Id & 0xFFFF_FFFF));
            w.WriteFixed32(2, account.TitleId);
        }
        else
        {
            w.WriteUInt64(1, account.Id);
            w.WriteUInt32(2, account.TitleId);
        }

        w.WriteUInt32(3, account.Region);
        return w.ToArray();
    }

    public static byte[] ChannelId(DiabloChannelId channel)
    {
        var host = new ProtoWriter();
        host.WriteUInt64(1, channel.HostLabel);
        host.WriteUInt64(2, channel.HostEpoch);
        var w = new ProtoWriter();
        w.WriteBytesField(2, host.ToArray());
        w.WriteFixed32(3, channel.Id);
        w.WriteUInt32(4, channel.Region);
        return w.ToArray();
    }

    /// <summary>ChannelService 23.</summary>
    public byte[] SendMessage(DiabloChannelId channel, string text)
    {
        if (string.IsNullOrEmpty(text) || System.Text.Encoding.UTF8.GetByteCount(text) > MaxMessageBytes)
        {
            throw new ArgumentException($"Chat text must be 1 to {MaxMessageBytes} UTF-8 bytes.", nameof(text));
        }

        var content = new ProtoWriter();
        content.WriteString(4, text);
        return Game == DiabloGame.D2R
            ? Fields((1, Handle()), (2, ChannelId(channel)), (3, content.ToArray()))
            : Fields((2, ChannelId(channel)), (3, content.ToArray()), (4, Handle()));
    }

    public byte[] UniqueType() => UniqueType(new UniqueChannelType(Program, PublicChannelType));

    public static byte[] UniqueType(UniqueChannelType type)
    {
        var w = new ProtoWriter();
        w.WriteFixed32(2, type.Program);
        w.WriteString(3, type.ChannelType);
        return w.ToArray();
    }

    /// <summary>ChannelService 5: the game's public channel list.</summary>
    public byte[] ListChannelTypes()
    {
        var wrapped = Fields((1, UniqueType()));
        return Game == DiabloGame.D2R
            ? Fields((1, Handle()), (2, wrapped))
            : Fields((2, wrapped), (4, Handle()));
    }

    /// <summary>ChannelService 6: find a public channel by the identity the list gave.</summary>
    public byte[] FindChannel(UniqueChannelType type, string identity)
    {
        var options = new ProtoWriter();
        options.WriteBytesField(1, UniqueType(type));
        options.WriteString(2, identity);
        options.WriteFixed32(3, EnUsLocale);
        return Game == DiabloGame.D2R
            ? Fields((1, Handle()), (2, options.ToArray()))
            : Fields((2, options.ToArray()), (3, Handle()));
    }

    /// <summary>ChannelService 10 (subscribe), 11 (unsubscribe) and 31 (leave) share this body.</summary>
    public byte[] ChannelAction(DiabloChannelId channel) =>
        Game == DiabloGame.D2R
            ? Fields((1, Handle()), (2, ChannelId(channel)))
            : Fields((2, ChannelId(channel)), (3, Handle()));

    /// <summary>AccountService 13: look up a BattleTag's account ID.</summary>
    public static byte[] ResolveBattleTag(string battleTag)
    {
        var tag = new ProtoWriter();
        tag.WriteString(4, battleTag);
        var w = new ProtoWriter();
        w.WriteBytesField(1, tag.ToArray());
        w.WriteUInt32(2, 1);
        return w.ToArray();
    }

    /// <summary>WhisperService 4.</summary>
    public static byte[] Whisper(ulong accountId, string text)
    {
        var content = new ProtoWriter();
        content.WriteString(1, text);
        var w = new ProtoWriter();
        w.WriteUInt64(1, accountId);
        w.WriteBytesField(2, content.ToArray());
        return w.ToArray();
    }

    /// <summary>
    /// The reply body to ConnectionService's echo (method 3): the request's
    /// fixed64 field 1 into field 1, its bytes field 3 into field 2.
    /// </summary>
    public static byte[] EchoReply(byte[] requestBody)
    {
        var request = ProtoFields.Parse(requestBody);
        var w = new ProtoWriter();
        if (request.Number(1) is { } time)
        {
            w.WriteFixed64(1, time);
        }

        w.WriteBytesField(2, request.Bytes(3));
        return w.ToArray();
    }

    private static byte[] Fields(params (int Field, byte[] Message)[] fields)
    {
        var w = new ProtoWriter();
        foreach (var (field, message) in fields)
        {
            w.WriteBytesField(field, message);
        }

        return w.ToArray();
    }
}
