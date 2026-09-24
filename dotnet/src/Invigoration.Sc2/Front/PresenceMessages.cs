using Invigoration.Sc2.Protobuf;

namespace Invigoration.Sc2.Front;

// bgs.protocol.presence.v1 messages. Field numbers come from TrinityCore's generated
// presence_types.pb.h / presence_service.pb.h / presence_listener.pb.h (11.0.2 regeneration),
// cross-checked against ncarrillo/superiority's prost output (bgs.protocol.presence.v1.rs) and
// HearthSim's hsproto (bnet/protocol/presence/presence.proto). The channel-listener delivery path
// (presence ChannelState as extension 101 of bgs.protocol.channel.v1.ChannelState) comes from
// TrinityCore's pre-2018-11 channel_types/presence_types and blizzless-diiis's RPCObject.cs.

/// <summary>bgs.protocol.presence.v1.FieldKey {program = 1, group = 2, field = 3, unique_id = 4}. Program is a plain uint32 FourCC here ("BN" = 0x424E), not fixed32.</summary>
public sealed class BgsPresenceFieldKey
{
    public required uint Program { get; init; }
    public required uint Group { get; init; }
    public required uint Field { get; init; }

    /// <summary>Distinguishes entries of a list-valued field (e.g. one per game account). Called "index" in older protos.</summary>
    public ulong? UniqueId { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteUInt32(1, Program);
        w.WriteUInt32(2, Group);
        w.WriteUInt32(3, Field);
        w.WriteUInt64(4, UniqueId);
        return w.ToArray();
    }

    public static BgsPresenceFieldKey Decode(byte[] data)
    {
        uint program = 0, group = 0, fieldNumber = 0;
        ulong? uniqueId = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: program = (uint)r.ReadVarint(); break;
                case 2: group = (uint)r.ReadVarint(); break;
                case 3: fieldNumber = (uint)r.ReadVarint(); break;
                case 4: uniqueId = r.ReadVarint(); break;
                default: r.Skip(type); break;
            }
        }

        return new BgsPresenceFieldKey { Program = program, Group = group, Field = fieldNumber, UniqueId = uniqueId };
    }
}

/// <summary>FieldOperation {field = 1 (Field {key = 1, value = 2}), operation = 2 (SET = 0, CLEAR = 1)}.</summary>
public sealed class BgsPresenceFieldOperation
{
    public required BgsPresenceFieldKey Key { get; init; }
    public required Variant Value { get; init; }
    public bool Clear { get; init; }

    public byte[] Encode()
    {
        var field = new ProtoWriter();
        field.WriteBytesField(1, Key.Encode());
        field.WriteBytesField(2, Value.Encode());

        var w = new ProtoWriter();
        w.WriteBytesField(1, field.ToArray());
        if (Clear) w.WriteUInt32(2, 1);
        return w.ToArray();
    }

    public static BgsPresenceFieldOperation Decode(byte[] data)
    {
        BgsPresenceFieldKey? key = null;
        Variant? value = null;
        var clear = false;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1:
                    var fr = new ProtoReader(r.ReadLengthDelimited());
                    while (fr.HasMore)
                    {
                        var (inner, innerType) = fr.ReadTag();
                        switch (inner)
                        {
                            case 1: key = BgsPresenceFieldKey.Decode(fr.ReadLengthDelimited()); break;
                            case 2: value = Variant.Decode(fr.ReadLengthDelimited()); break;
                            default: fr.Skip(innerType); break;
                        }
                    }

                    break;
                case 2: clear = r.ReadVarint() == 1; break;
                default: r.Skip(type); break;
            }
        }

        return new BgsPresenceFieldOperation
        {
            Key = key ?? throw new InvalidOperationException("Presence field has no key."),
            Value = value ?? new Variant(),
            Clear = clear,
        };
    }
}

/// <summary>PresenceState {entity_id = 1, field_operation = 2} and presence ChannelState, which adds healing = 3.</summary>
public sealed class BgsPresenceState
{
    public EntityId? EntityId { get; init; }
    public List<BgsPresenceFieldOperation> Operations { get; init; } = [];
    public bool? Healing { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        if (EntityId is not null) w.WriteBytesField(1, EntityId.Encode());
        foreach (var op in Operations) w.WriteBytesField(2, op.Encode());
        w.WriteBool(3, Healing);
        return w.ToArray();
    }

    public static BgsPresenceState Decode(byte[] data)
    {
        EntityId? entity = null;
        List<BgsPresenceFieldOperation> ops = [];
        bool? healing = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: entity = EntityId.Decode(r.ReadLengthDelimited()); break;
                case 2: ops.Add(BgsPresenceFieldOperation.Decode(r.ReadLengthDelimited())); break;
                case 3: healing = r.ReadVarint() != 0; break;
                default: r.Skip(type); break;
            }
        }

        return new BgsPresenceState { EntityId = entity, Operations = ops, Healing = healing };
    }
}

/// <summary>PresenceService.Subscribe (method 1): {agent_id = 1, entity_id = 2, object_id = 3, program = 4 (repeated fixed32), key = 6}.</summary>
public sealed class BgsPresenceSubscribeRequest
{
    public EntityId? AgentId { get; init; }
    public required EntityId EntityId { get; init; }
    public required ulong ObjectId { get; init; }
    public List<uint> Programs { get; init; } = [];
    public List<BgsPresenceFieldKey> Keys { get; init; } = [];

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        if (AgentId is not null) w.WriteBytesField(1, AgentId.Encode());
        w.WriteBytesField(2, EntityId.Encode());
        w.WriteUInt64(3, ObjectId);
        foreach (var p in Programs) w.WriteFixed32(4, p);
        foreach (var k in Keys) w.WriteBytesField(6, k.Encode());
        return w.ToArray();
    }

    public static BgsPresenceSubscribeRequest Decode(byte[] data)
    {
        EntityId? agent = null, entity = null;
        ulong objectId = 0;
        List<uint> programs = [];
        List<BgsPresenceFieldKey> keys = [];
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: agent = EntityId.Decode(r.ReadLengthDelimited()); break;
                case 2: entity = EntityId.Decode(r.ReadLengthDelimited()); break;
                case 3: objectId = r.ReadVarint(); break;
                case 4: ProtoPacking.ReadFixed32s(r, type, programs); break;
                case 6: keys.Add(BgsPresenceFieldKey.Decode(r.ReadLengthDelimited())); break;
                default: r.Skip(type); break;
            }
        }

        return new BgsPresenceSubscribeRequest { AgentId = agent, EntityId = entity!, ObjectId = objectId, Programs = programs, Keys = keys };
    }
}

/// <summary>PresenceService.Unsubscribe (method 2): {agent_id = 1, entity_id = 2, object_id = 3}.</summary>
public sealed class BgsPresenceUnsubscribeRequest
{
    public EntityId? AgentId { get; init; }
    public required EntityId EntityId { get; init; }
    public ulong? ObjectId { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        if (AgentId is not null) w.WriteBytesField(1, AgentId.Encode());
        w.WriteBytesField(2, EntityId.Encode());
        w.WriteUInt64(3, ObjectId);
        return w.ToArray();
    }
}

/// <summary>PresenceService.BatchSubscribe (method 8): {agent_id = 1, entity_id = 2 (repeated), program = 3, key = 4, object_id = 5}.</summary>
public sealed class BgsPresenceBatchSubscribeRequest
{
    public EntityId? AgentId { get; init; }
    public List<EntityId> EntityIds { get; init; } = [];
    public List<uint> Programs { get; init; } = [];
    public List<BgsPresenceFieldKey> Keys { get; init; } = [];
    public ulong? ObjectId { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        if (AgentId is not null) w.WriteBytesField(1, AgentId.Encode());
        foreach (var e in EntityIds) w.WriteBytesField(2, e.Encode());
        foreach (var p in Programs) w.WriteFixed32(3, p);
        foreach (var k in Keys) w.WriteBytesField(4, k.Encode());
        w.WriteUInt64(5, ObjectId);
        return w.ToArray();
    }

    public static BgsPresenceBatchSubscribeRequest Decode(byte[] data)
    {
        EntityId? agent = null;
        List<EntityId> entities = [];
        List<uint> programs = [];
        List<BgsPresenceFieldKey> keys = [];
        ulong? objectId = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: agent = EntityId.Decode(r.ReadLengthDelimited()); break;
                case 2: entities.Add(EntityId.Decode(r.ReadLengthDelimited())); break;
                case 3: ProtoPacking.ReadFixed32s(r, type, programs); break;
                case 4: keys.Add(BgsPresenceFieldKey.Decode(r.ReadLengthDelimited())); break;
                case 5: objectId = r.ReadVarint(); break;
                default: r.Skip(type); break;
            }
        }

        return new BgsPresenceBatchSubscribeRequest { AgentId = agent, EntityIds = entities, Programs = programs, Keys = keys, ObjectId = objectId };
    }
}

/// <summary>BatchSubscribeResponse {subscribe_failed = 1: repeated SubscribeResult {entity_id = 1, result = 2}}.</summary>
public sealed class BgsPresenceBatchSubscribeResponse
{
    public List<(EntityId? EntityId, uint? Result)> Failed { get; init; } = [];

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        foreach (var (entity, result) in Failed)
        {
            var inner = new ProtoWriter();
            if (entity is not null) inner.WriteBytesField(1, entity.Encode());
            inner.WriteUInt32(2, result);
            w.WriteBytesField(1, inner.ToArray());
        }

        return w.ToArray();
    }

    public static BgsPresenceBatchSubscribeResponse Decode(byte[] data)
    {
        List<(EntityId?, uint?)> failed = [];
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            if (field != 1)
            {
                r.Skip(type);
                continue;
            }

            EntityId? entity = null;
            uint? result = null;
            var ir = new ProtoReader(r.ReadLengthDelimited());
            while (ir.HasMore)
            {
                var (inner, innerType) = ir.ReadTag();
                switch (inner)
                {
                    case 1: entity = EntityId.Decode(ir.ReadLengthDelimited()); break;
                    case 2: result = (uint)ir.ReadVarint(); break;
                    default: ir.Skip(innerType); break;
                }
            }

            failed.Add((entity, result));
        }

        return new BgsPresenceBatchSubscribeResponse { Failed = failed };
    }
}

/// <summary>
/// PresenceListener ("bnet.protocol.presence.v1.PresenceListener") OnSubscribe (method 1,
/// SubscribeNotification) and OnStateChanged (method 2, StateChangedNotification). Both have
/// the same shape: {subscriber_id = 1 (account.v1.AccountId), state = 2 (repeated PresenceState),
/// subscriber_program = 3}. Both are NO_RESPONSE.
/// </summary>
public sealed class BgsPresenceListenerNotification
{
    public List<BgsPresenceState> States { get; init; } = [];
    public uint? SubscriberProgram { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        foreach (var s in States) w.WriteBytesField(2, s.Encode());
        w.WriteUInt32(3, SubscriberProgram);
        return w.ToArray();
    }

    public static BgsPresenceListenerNotification Decode(byte[] data)
    {
        List<BgsPresenceState> states = [];
        uint? program = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 2: states.Add(BgsPresenceState.Decode(r.ReadLengthDelimited())); break;
                case 3 when type == WireType.Varint: program = (uint)r.ReadVarint(); break;
                default: r.Skip(type); break;
            }
        }

        return new BgsPresenceListenerNotification { States = states, SubscriberProgram = program };
    }
}

/// <summary>
/// The older delivery path: PresenceService.Subscribe answered through ChannelListener
/// ("bnet.protocol.channel.ChannelSubscriber"), OnJoin (method 1, JoinNotification, channel_state = 3)
/// for the initial state and OnUpdateChannelState (method 6, UpdateChannelStateNotification,
/// state_change = 2) for changes. The presence fields sit in extension 101 of
/// bgs.protocol.channel.v1.ChannelState. Header.object_id carries the Subscribe's object_id.
/// </summary>
public static class BgsChannelPresence
{
    public const uint OnJoinMethod = 1;
    public const uint OnUpdateChannelStateMethod = 6;
    public const int PresenceExtensionField = 101;

    /// <summary>Returns the presence state inside a ChannelListener call, or null when the call carries none.</summary>
    public static BgsPresenceState? Decode(uint methodId, byte[] data)
    {
        var channelStateField = methodId switch
        {
            OnJoinMethod => 3,
            OnUpdateChannelStateMethod => 2,
            _ => 0,
        };
        if (channelStateField == 0)
        {
            return null;
        }

        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            if (field != channelStateField || type != WireType.LengthDelimited)
            {
                r.Skip(type);
                continue;
            }

            var cr = new ProtoReader(r.ReadLengthDelimited());
            while (cr.HasMore)
            {
                var (inner, innerType) = cr.ReadTag();
                if (inner == PresenceExtensionField && innerType == WireType.LengthDelimited)
                {
                    return BgsPresenceState.Decode(cr.ReadLengthDelimited());
                }

                cr.Skip(innerType);
            }
        }

        return null;
    }

    /// <summary>Builds an UpdateChannelStateNotification (or JoinNotification) body carrying <paramref name="state"/>; for tests.</summary>
    public static byte[] Encode(uint methodId, BgsPresenceState state)
    {
        var channelState = new ProtoWriter();
        channelState.WriteBytesField(PresenceExtensionField, state.Encode());
        var w = new ProtoWriter();
        w.WriteBytesField(methodId == OnJoinMethod ? 3 : 2, channelState.ToArray());
        return w.ToArray();
    }
}
