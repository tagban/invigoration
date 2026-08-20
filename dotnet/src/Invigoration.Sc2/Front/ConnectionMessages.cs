using Invigoration.Sc2.Protobuf;

namespace Invigoration.Sc2.Front;

/// <summary>bgs.protocol.connection.v1 messages — the anonymous bind/connect service (service_id 0).</summary>
public sealed class ConnectRequest
{
    public ProcessId? ClientId { get; init; }
    public BindRequest? BindRequest { get; init; }
    public bool? UseBindlessRpc { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        if (ClientId is not null) w.WriteBytesField(1, ClientId.Encode());
        if (BindRequest is not null) w.WriteBytesField(2, BindRequest.Encode());
        w.WriteBool(3, UseBindlessRpc);
        return w.ToArray();
    }

    public static ConnectRequest Decode(byte[] data)
    {
        ProcessId? clientId = null;
        BindRequest? bindRequest = null;
        bool? useBindlessRpc = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: clientId = ProcessId.Decode(r.ReadLengthDelimited()); break;
                case 2: bindRequest = Front.BindRequest.Decode(r.ReadLengthDelimited()); break;
                case 3: useBindlessRpc = r.ReadVarint() != 0; break;
                default: r.Skip(type); break;
            }
        }

        return new ConnectRequest { ClientId = clientId, BindRequest = bindRequest, UseBindlessRpc = useBindlessRpc };
    }
}

public sealed class ConnectResponse
{
    public required ProcessId ServerId { get; init; }
    public ProcessId? ClientId { get; init; }
    public uint? BindResult { get; init; }
    public BindResponse? BindResponse { get; init; }
    public ulong? ServerTime { get; init; }
    public bool? UseBindlessRpc { get; init; }

    public static ConnectResponse Decode(byte[] data)
    {
        ProcessId? serverId = null, clientId = null;
        uint? bindResult = null;
        BindResponse? bindResponse = null;
        ulong? serverTime = null;
        bool? useBindlessRpc = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: serverId = ProcessId.Decode(r.ReadLengthDelimited()); break;
                case 2: clientId = ProcessId.Decode(r.ReadLengthDelimited()); break;
                case 3: bindResult = (uint)r.ReadVarint(); break;
                case 4: bindResponse = Front.BindResponse.Decode(r.ReadLengthDelimited()); break;
                case 6: serverTime = r.ReadVarint(); break;
                case 7: useBindlessRpc = r.ReadVarint() != 0; break;
                default: r.Skip(type); break;
            }
        }

        return new ConnectResponse
        {
            ServerId = serverId!,
            ClientId = clientId,
            BindResult = bindResult,
            BindResponse = bindResponse,
            ServerTime = serverTime,
            UseBindlessRpc = useBindlessRpc,
        };
    }
}

public sealed class BoundService
{
    public required uint Hash { get; init; }
    public required uint Id { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteFixed32(1, Hash);
        w.WriteUInt32(2, Id);
        return w.ToArray();
    }

    public static BoundService Decode(byte[] data)
    {
        uint hash = 0, id = 0;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: hash = r.ReadFixed32(); break;
                case 2: id = (uint)r.ReadVarint(); break;
                default: r.Skip(type); break;
            }
        }

        return new BoundService { Hash = hash, Id = id };
    }
}

public sealed class BindRequest
{
    public List<BoundService> ExportedServices { get; init; } = [];
    public List<BoundService> ImportedServices { get; init; } = [];

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        foreach (var s in ExportedServices) w.WriteBytesField(3, s.Encode());
        foreach (var s in ImportedServices) w.WriteBytesField(4, s.Encode());
        return w.ToArray();
    }

    public static BindRequest Decode(byte[] data)
    {
        List<BoundService> exported = [], imported = [];
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 3: exported.Add(BoundService.Decode(r.ReadLengthDelimited())); break;
                case 4: imported.Add(BoundService.Decode(r.ReadLengthDelimited())); break;
                default: r.Skip(type); break;
            }
        }

        return new BindRequest { ExportedServices = exported, ImportedServices = imported };
    }
}

public sealed class BindResponse
{
    public List<uint> ImportedServiceIds { get; init; } = [];

    public static BindResponse Decode(byte[] data)
    {
        List<uint> ids = [];
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            if (field == 1)
            {
                ids.Add((uint)r.ReadVarint());
            }
            else
            {
                r.Skip(type);
            }
        }

        return new BindResponse { ImportedServiceIds = ids };
    }
}

public sealed class EchoRequest
{
    public ulong? Time { get; init; }
    public bool? NetworkOnly { get; init; }
    public byte[]? Payload { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteFixed64(1, Time);
        w.WriteBool(2, NetworkOnly);
        w.WriteBytesField(3, Payload);
        return w.ToArray();
    }

    public static EchoRequest Decode(byte[] data)
    {
        ulong? time = null;
        bool? networkOnly = null;
        byte[]? payload = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: time = r.ReadFixed64(); break;
                case 2: networkOnly = r.ReadVarint() != 0; break;
                case 3: payload = r.ReadLengthDelimited(); break;
                default: r.Skip(type); break;
            }
        }

        return new EchoRequest { Time = time, NetworkOnly = networkOnly, Payload = payload };
    }
}

public sealed class EchoResponse
{
    public ulong? Time { get; init; }
    public byte[]? Payload { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteFixed64(1, Time);
        w.WriteBytesField(2, Payload);
        return w.ToArray();
    }

    public static EchoResponse Decode(byte[] data)
    {
        ulong? time = null;
        byte[]? payload = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: time = r.ReadFixed64(); break;
                case 2: payload = r.ReadLengthDelimited(); break;
                default: r.Skip(type); break;
            }
        }

        return new EchoResponse { Time = time, Payload = payload };
    }
}

public sealed class DisconnectNotification
{
    public required uint ErrorCode { get; init; }
    public string? Reason { get; init; }

    public static DisconnectNotification Decode(byte[] data)
    {
        uint errorCode = 0;
        string? reason = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: errorCode = (uint)r.ReadVarint(); break;
                case 2: reason = r.ReadString(); break;
                default: r.Skip(type); break;
            }
        }

        return new DisconnectNotification { ErrorCode = errorCode, Reason = reason };
    }
}
