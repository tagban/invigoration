using Invigoration.Sc2.Protobuf;

namespace Invigoration.Sc2.Front;

/// <summary>bgs.protocol.game_utilities.v1 messages — used to hand off from Front to the native ("Sunken") layer.</summary>
public sealed class ClientInfo
{
    public string? ClientAddress { get; init; }
    public bool? PrivilegedNetwork { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteString(1, ClientAddress);
        w.WriteBool(2, PrivilegedNetwork);
        return w.ToArray();
    }
}

public sealed class ClientRequest
{
    public List<Attribute> Attributes { get; init; } = [];
    public EntityId? AccountId { get; init; }
    public EntityId? GameAccountId { get; init; }
    public uint? Program { get; init; }
    public ClientInfo? ClientInfo { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        foreach (var a in Attributes) w.WriteBytesField(1, a.Encode());
        if (AccountId is not null) w.WriteBytesField(3, AccountId.Encode());
        if (GameAccountId is not null) w.WriteBytesField(4, GameAccountId.Encode());
        w.WriteFixed32(5, Program);
        if (ClientInfo is not null) w.WriteBytesField(6, ClientInfo.Encode());
        return w.ToArray();
    }

    public static ClientRequest Decode(byte[] data)
    {
        List<Attribute> attributes = [];
        EntityId? accountId = null;
        EntityId? gameAccountId = null;
        uint? program = null;

        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: attributes.Add(Attribute.Decode(r.ReadLengthDelimited())); break;
                case 3: accountId = EntityId.Decode(r.ReadLengthDelimited()); break;
                case 4: gameAccountId = EntityId.Decode(r.ReadLengthDelimited()); break;
                case 5: program = r.ReadFixed32(); break;
                case 6: r.ReadLengthDelimited(); break; // ClientInfo, not needed by this port
                default: r.Skip(type); break;
            }
        }

        return new ClientRequest { Attributes = attributes, AccountId = accountId, GameAccountId = gameAccountId, Program = program };
    }
}

public sealed class ClientResponse
{
    public List<Attribute> Attributes { get; init; } = [];

    public static ClientResponse Decode(byte[] data)
    {
        List<Attribute> attributes = [];
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            if (field == 1)
            {
                attributes.Add(Attribute.Decode(r.ReadLengthDelimited()));
            }
            else
            {
                r.Skip(type);
            }
        }

        return new ClientResponse { Attributes = attributes };
    }
}
