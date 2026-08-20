using Invigoration.Sc2.Protobuf;

namespace Invigoration.Sc2.Front;

/// <summary>
/// Shared protobuf types from bgs.protocol.rs, reconstructed from the
/// generated Rust prost output (no .proto source exists upstream). Field
/// numbers and wire types below are copied verbatim from that file's
/// #[prost(...)] attributes.
/// </summary>
public sealed class ProcessId
{
    public required uint Label { get; init; }
    public required uint Epoch { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteUInt32(1, Label);
        w.WriteUInt32(2, Epoch);
        return w.ToArray();
    }

    public static ProcessId Decode(byte[] data)
    {
        uint label = 0, epoch = 0;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: label = (uint)r.ReadVarint(); break;
                case 2: epoch = (uint)r.ReadVarint(); break;
                default: r.Skip(type); break;
            }
        }

        return new ProcessId { Label = label, Epoch = epoch };
    }
}

public sealed class ObjectAddress
{
    public required ProcessId Host { get; init; }
    public ulong? ObjectId { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteBytesField(1, Host.Encode());
        w.WriteUInt64(2, ObjectId);
        return w.ToArray();
    }

    public static ObjectAddress Decode(byte[] data)
    {
        ProcessId? host = null;
        ulong? objectId = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: host = ProcessId.Decode(r.ReadLengthDelimited()); break;
                case 2: objectId = r.ReadVarint(); break;
                default: r.Skip(type); break;
            }
        }

        return new ObjectAddress { Host = host!, ObjectId = objectId };
    }
}

public sealed class ErrorInfo
{
    public required ObjectAddress ObjectAddress { get; init; }
    public required uint Status { get; init; }
    public required uint ServiceHash { get; init; }
    public required uint MethodId { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteBytesField(1, ObjectAddress.Encode());
        w.WriteUInt32(2, Status);
        w.WriteUInt32(3, ServiceHash);
        w.WriteUInt32(4, MethodId);
        return w.ToArray();
    }

    public static ErrorInfo Decode(byte[] data)
    {
        ObjectAddress? objectAddress = null;
        uint status = 0, serviceHash = 0, methodId = 0;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: objectAddress = ObjectAddress.Decode(r.ReadLengthDelimited()); break;
                case 2: status = (uint)r.ReadVarint(); break;
                case 3: serviceHash = (uint)r.ReadVarint(); break;
                case 4: methodId = (uint)r.ReadVarint(); break;
                default: r.Skip(type); break;
            }
        }

        return new ErrorInfo { ObjectAddress = objectAddress!, Status = status, ServiceHash = serviceHash, MethodId = methodId };
    }
}

public sealed class FanoutTarget
{
    public string? ClientId { get; init; }
    public byte[]? Key { get; init; }
    public ulong? ObjectId { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteString(1, ClientId);
        w.WriteBytesField(2, Key);
        w.WriteUInt64(3, ObjectId);
        return w.ToArray();
    }

    public static FanoutTarget Decode(byte[] data)
    {
        string? clientId = null;
        byte[]? key = null;
        ulong? objectId = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: clientId = r.ReadString(); break;
                case 2: key = r.ReadLengthDelimited(); break;
                case 3: objectId = r.ReadVarint(); break;
                default: r.Skip(type); break;
            }
        }

        return new FanoutTarget { ClientId = clientId, Key = key, ObjectId = objectId };
    }
}

public sealed class EntityId
{
    public required ulong High { get; init; }
    public required ulong Low { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteFixed64(1, High);
        w.WriteFixed64(2, Low);
        return w.ToArray();
    }

    public static EntityId Decode(byte[] data)
    {
        ulong high = 0, low = 0;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: high = r.ReadFixed64(); break;
                case 2: low = r.ReadFixed64(); break;
                default: r.Skip(type); break;
            }
        }

        return new EntityId { High = high, Low = low };
    }
}

public sealed class Identity
{
    public EntityId? AccountId { get; init; }
    public EntityId? GameAccountId { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        if (AccountId is not null) w.WriteBytesField(1, AccountId.Encode());
        if (GameAccountId is not null) w.WriteBytesField(2, GameAccountId.Encode());
        return w.ToArray();
    }

    public static Identity Decode(byte[] data)
    {
        EntityId? accountId = null, gameAccountId = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: accountId = EntityId.Decode(r.ReadLengthDelimited()); break;
                case 2: gameAccountId = EntityId.Decode(r.ReadLengthDelimited()); break;
                default: r.Skip(type); break;
            }
        }

        return new Identity { AccountId = accountId, GameAccountId = gameAccountId };
    }
}

public sealed class Variant
{
    public bool? BoolValue { get; init; }
    public long? IntValue { get; init; }
    public double? FloatValue { get; init; }
    public string? StringValue { get; init; }
    public byte[]? BlobValue { get; init; }
    public byte[]? MessageValue { get; init; }
    public string? FourccValue { get; init; }
    public ulong? UintValue { get; init; }
    public EntityId? EntityIdValue { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteBool(2, BoolValue);
        w.WriteInt64(3, IntValue);
        w.WriteDouble(4, FloatValue);
        w.WriteString(5, StringValue);
        w.WriteBytesField(6, BlobValue);
        w.WriteBytesField(7, MessageValue);
        w.WriteString(8, FourccValue);
        w.WriteUInt64(9, UintValue);
        if (EntityIdValue is not null) w.WriteBytesField(10, EntityIdValue.Encode());
        return w.ToArray();
    }

    public static Variant Decode(byte[] data)
    {
        bool? boolValue = null;
        long? intValue = null;
        double? floatValue = null;
        string? stringValue = null, fourccValue = null;
        byte[]? blobValue = null, messageValue = null;
        ulong? uintValue = null;
        EntityId? entityIdValue = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 2: boolValue = r.ReadVarint() != 0; break;
                case 3: intValue = unchecked((long)r.ReadVarint()); break;
                case 4: floatValue = r.ReadDouble(); break;
                case 5: stringValue = r.ReadString(); break;
                case 6: blobValue = r.ReadLengthDelimited(); break;
                case 7: messageValue = r.ReadLengthDelimited(); break;
                case 8: fourccValue = r.ReadString(); break;
                case 9: uintValue = r.ReadVarint(); break;
                case 10: entityIdValue = EntityId.Decode(r.ReadLengthDelimited()); break;
                default: r.Skip(type); break;
            }
        }

        return new Variant
        {
            BoolValue = boolValue,
            IntValue = intValue,
            FloatValue = floatValue,
            StringValue = stringValue,
            BlobValue = blobValue,
            MessageValue = messageValue,
            FourccValue = fourccValue,
            UintValue = uintValue,
            EntityIdValue = entityIdValue,
        };
    }
}

public sealed class Attribute
{
    public required string Name { get; init; }
    public required Variant Value { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteString(1, Name);
        w.WriteBytesField(2, Value.Encode());
        return w.ToArray();
    }

    public static Attribute Decode(byte[] data)
    {
        string? name = null;
        Variant? value = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: name = r.ReadString(); break;
                case 2: value = Variant.Decode(r.ReadLengthDelimited()); break;
                default: r.Skip(type); break;
            }
        }

        return new Attribute { Name = name!, Value = value! };
    }
}

public sealed class ContentHandle
{
    public required uint Region { get; init; }
    public required uint Usage { get; init; }
    public required byte[] Hash { get; init; }
    public string? ProtoUrl { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteFixed32(1, Region);
        w.WriteFixed32(2, Usage);
        w.WriteBytesField(3, Hash);
        w.WriteString(4, ProtoUrl);
        return w.ToArray();
    }

    public static ContentHandle Decode(byte[] data)
    {
        uint region = 0, usage = 0;
        byte[]? hash = null;
        string? protoUrl = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: region = r.ReadFixed32(); break;
                case 2: usage = r.ReadFixed32(); break;
                case 3: hash = r.ReadLengthDelimited(); break;
                case 4: protoUrl = r.ReadString(); break;
                default: r.Skip(type); break;
            }
        }

        return new ContentHandle { Region = region, Usage = usage, Hash = hash!, ProtoUrl = protoUrl };
    }
}
