using Invigoration.Sc2.Protobuf;

namespace Invigoration.Sc2.Front;

// bgs.protocol.friends.v1 messages: the Battle.net ACCOUNT friends list (the one the Battle.net
// app shows), not the SC2-native Sunken friends records.
//
// Field numbers come from TrinityCore's generated friends_types.pb.h / friends_service.pb.h
// (src/server/proto/Client, 11.0.2 regeneration, 2024-07-26). Where Blizzard has moved fields
// between SDK versions, the decoders accept both layouts and tell them apart by wire type:
// HearthSim's hsproto (bnet/protocol/friends/friends.proto, 2019) carries Friend.full_name = 6
// and Friend.battle_tag = 7 as strings, while current TrinityCore has Friend.creation_time = 6
// (uint64) and no names on Friend at all (names then come from presence, see BgsPresenceFields).

/// <summary>FriendsService.Subscribe (method 1) and Unsubscribe (method 11) request: {agent_id = 1, object_id = 2}.</summary>
public sealed class BgsFriendsSubscribeRequest
{
    public EntityId? AgentId { get; init; }

    /// <summary>Our listener id. Required on Subscribe; later FriendsListener calls may carry it as Header.object_id.</summary>
    public ulong? ObjectId { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        if (AgentId is not null) w.WriteBytesField(1, AgentId.Encode());
        w.WriteUInt64(2, ObjectId);
        return w.ToArray();
    }

    public static BgsFriendsSubscribeRequest Decode(byte[] data)
    {
        EntityId? agent = null;
        ulong? objectId = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: agent = EntityId.Decode(r.ReadLengthDelimited()); break;
                case 2: objectId = r.ReadVarint(); break;
                default: r.Skip(type); break;
            }
        }

        return new BgsFriendsSubscribeRequest { AgentId = agent, ObjectId = objectId };
    }
}

/// <summary>bgs.protocol.friends.v1.Friend.</summary>
public sealed class BgsFriendMessage
{
    /// <summary>The friend's Battle.net ACCOUNT entity id (field 1). Subscribe presence to this.</summary>
    public required EntityId AccountId { get; init; }
    public List<Attribute> Attributes { get; init; } = [];
    public List<uint> Roles { get; init; } = [];
    public ulong? Privileges { get; init; }
    public ulong? AttributesEpoch { get; init; }

    /// <summary>Legacy layout only (field 6 as a string). Null on servers that send the current layout.</summary>
    public string? FullName { get; init; }

    /// <summary>Legacy layout only (field 7). Null on servers that send the current layout.</summary>
    public string? BattleTag { get; init; }

    /// <summary>Current layout only (field 6 as a varint).</summary>
    public ulong? CreationTime { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteBytesField(1, AccountId.Encode());
        foreach (var a in Attributes) w.WriteBytesField(2, a.Encode());
        ProtoPacking.WritePackedVarints(w, 3, Roles);
        w.WriteUInt64(4, Privileges);
        w.WriteUInt64(5, AttributesEpoch);
        if (FullName is not null)
        {
            w.WriteString(6, FullName);
        }
        else
        {
            w.WriteUInt64(6, CreationTime);
        }

        w.WriteString(7, BattleTag);
        return w.ToArray();
    }

    public static BgsFriendMessage Decode(byte[] data)
    {
        EntityId? accountId = null;
        List<Attribute> attributes = [];
        List<uint> roles = [];
        ulong? privileges = null, epoch = null, creation = null;
        string? fullName = null, battleTag = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: accountId = EntityId.Decode(r.ReadLengthDelimited()); break;
                case 2: attributes.Add(Attribute.Decode(r.ReadLengthDelimited())); break;
                case 3: ProtoPacking.ReadVarints(r, type, roles); break;
                case 4: privileges = r.ReadVarint(); break;
                case 5: epoch = r.ReadVarint(); break;
                case 6 when type == WireType.LengthDelimited: fullName = r.ReadString(); break;
                case 6 when type == WireType.Varint: creation = r.ReadVarint(); break;
                case 7 when type == WireType.LengthDelimited: battleTag = r.ReadString(); break;
                default: r.Skip(type); break;
            }
        }

        return new BgsFriendMessage
        {
            AccountId = accountId ?? throw new InvalidOperationException("Friend has no account_id."),
            Attributes = attributes,
            Roles = roles,
            Privileges = privileges,
            AttributesEpoch = epoch,
            FullName = fullName,
            BattleTag = battleTag,
            CreationTime = creation,
        };
    }
}

/// <summary>bgs.protocol.friends.v1.ReceivedInvitation: a friend request someone sent us.</summary>
public sealed class BgsReceivedInvitation
{
    public required ulong Id { get; init; }
    public Identity? InviterIdentity { get; init; }
    public Identity? InviteeIdentity { get; init; }
    public string? InviterName { get; init; }
    public string? InviteeName { get; init; }
    public ulong? CreationTime { get; init; }
    public uint? Program { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteFixed64(1, Id);
        if (InviterIdentity is not null) w.WriteBytesField(2, InviterIdentity.Encode());
        if (InviteeIdentity is not null) w.WriteBytesField(3, InviteeIdentity.Encode());
        w.WriteString(4, InviterName);
        w.WriteString(5, InviteeName);
        w.WriteUInt64(7, CreationTime);
        w.WriteFixed32(9, Program);
        return w.ToArray();
    }

    public static BgsReceivedInvitation Decode(byte[] data)
    {
        ulong id = 0;
        Identity? inviter = null, invitee = null;
        string? inviterName = null, inviteeName = null;
        ulong? creation = null;
        uint? program = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: id = r.ReadFixed64(); break;
                case 2: inviter = Identity.Decode(r.ReadLengthDelimited()); break;
                case 3: invitee = Identity.Decode(r.ReadLengthDelimited()); break;
                case 4: inviterName = r.ReadString(); break;
                case 5: inviteeName = r.ReadString(); break;
                case 7: creation = r.ReadVarint(); break;
                case 9 when type == WireType.Fixed32: program = r.ReadFixed32(); break;
                default: r.Skip(type); break;
            }
        }

        return new BgsReceivedInvitation
        {
            Id = id,
            InviterIdentity = inviter,
            InviteeIdentity = invitee,
            InviterName = inviterName,
            InviteeName = inviteeName,
            CreationTime = creation,
            Program = program,
        };
    }
}

/// <summary>bgs.protocol.friends.v1.SentInvitation: a friend request we sent.</summary>
public sealed class BgsSentInvitation
{
    public ulong? Id { get; init; }
    public string? TargetName { get; init; }
    public uint? Role { get; init; }
    public ulong? CreationTime { get; init; }
    public uint? Program { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteFixed64(1, Id);
        w.WriteString(2, TargetName);
        w.WriteUInt32(3, Role);
        w.WriteUInt64(5, CreationTime);
        w.WriteFixed32(6, Program);
        return w.ToArray();
    }

    public static BgsSentInvitation Decode(byte[] data)
    {
        ulong? id = null, creation = null;
        string? target = null;
        uint? role = null, program = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: id = r.ReadFixed64(); break;
                case 2: target = r.ReadString(); break;
                case 3: role = (uint)r.ReadVarint(); break;
                case 5: creation = r.ReadVarint(); break;
                case 6 when type == WireType.Fixed32: program = r.ReadFixed32(); break;
                default: r.Skip(type); break;
            }
        }

        return new BgsSentInvitation { Id = id, TargetName = target, Role = role, CreationTime = creation, Program = program };
    }
}

/// <summary>bgs.protocol.friends.v1.SubscribeResponse: the full friends list as of subscribing.</summary>
public sealed class BgsFriendsSubscribeResponse
{
    public uint? MaxFriends { get; init; }
    public uint? MaxReceivedInvitations { get; init; }
    public uint? MaxSentInvitations { get; init; }
    public List<BgsFriendMessage> Friends { get; init; } = [];
    public List<BgsReceivedInvitation> ReceivedInvitations { get; init; } = [];
    public List<BgsSentInvitation> SentInvitations { get; init; } = [];

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteUInt32(1, MaxFriends);
        w.WriteUInt32(2, MaxReceivedInvitations);
        w.WriteUInt32(3, MaxSentInvitations);
        foreach (var f in Friends) w.WriteBytesField(5, f.Encode());
        foreach (var i in ReceivedInvitations) w.WriteBytesField(7, i.Encode());
        foreach (var i in SentInvitations) w.WriteBytesField(8, i.Encode());
        return w.ToArray();
    }

    public static BgsFriendsSubscribeResponse Decode(byte[] data)
    {
        uint? maxFriends = null, maxReceived = null, maxSent = null;
        List<BgsFriendMessage> friends = [];
        List<BgsReceivedInvitation> received = [];
        List<BgsSentInvitation> sent = [];
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: maxFriends = (uint)r.ReadVarint(); break;
                case 2: maxReceived = (uint)r.ReadVarint(); break;
                case 3: maxSent = (uint)r.ReadVarint(); break;
                case 5: friends.Add(BgsFriendMessage.Decode(r.ReadLengthDelimited())); break;
                case 7: received.Add(BgsReceivedInvitation.Decode(r.ReadLengthDelimited())); break;
                case 8: sent.Add(BgsSentInvitation.Decode(r.ReadLengthDelimited())); break;
                default: r.Skip(type); break; // 4 = repeated bgs.protocol.Role definitions, not needed
            }
        }

        return new BgsFriendsSubscribeResponse
        {
            MaxFriends = maxFriends,
            MaxReceivedInvitations = maxReceived,
            MaxSentInvitations = maxSent,
            Friends = friends,
            ReceivedInvitations = received,
            SentInvitations = sent,
        };
    }
}

/// <summary>What a FriendsListener callback says happened.</summary>
public enum BgsFriendsNotificationKind
{
    FriendAdded = 1,
    FriendRemoved = 2,
    ReceivedInvitationAdded = 3,
    ReceivedInvitationRemoved = 4,
    SentInvitationAdded = 5,
    SentInvitationRemoved = 6,
    FriendUpdated = 7,
}

/// <summary>
/// One decoded FriendsListener ("bnet.protocol.friends.FriendsNotify") callback. The kind equals
/// the listener method id. Bodies: FriendNotification {target = 1, account_id = 5} for 1/2,
/// UpdateFriendStateNotification {changed_friend = 1, account_id = 5} for 7,
/// InvitationNotification {invitation = 1, reason = 3, account_id = 5} for 3/4,
/// SentInvitationAddedNotification {account_id = 1, invitation = 2} for 5 and
/// SentInvitationRemovedNotification {account_id = 1, invitation_id = 2, reason = 3} for 6.
/// All seven are NO_RESPONSE, so the client sends nothing back.
/// </summary>
public sealed class BgsFriendsNotification
{
    public required BgsFriendsNotificationKind Kind { get; init; }
    public BgsFriendMessage? Friend { get; init; }
    public BgsReceivedInvitation? ReceivedInvitation { get; init; }
    public BgsSentInvitation? SentInvitation { get; init; }
    public ulong? InvitationId { get; init; }
    public uint? Reason { get; init; }

    /// <summary>The subscriber's own account (field 5, or field 1 for the sent-invitation callbacks).</summary>
    public EntityId? SubscriberAccountId { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        switch (Kind)
        {
            case BgsFriendsNotificationKind.FriendAdded or BgsFriendsNotificationKind.FriendRemoved or BgsFriendsNotificationKind.FriendUpdated:
                w.WriteBytesField(1, Friend?.Encode());
                if (SubscriberAccountId is not null) w.WriteBytesField(5, SubscriberAccountId.Encode());
                break;
            case BgsFriendsNotificationKind.ReceivedInvitationAdded or BgsFriendsNotificationKind.ReceivedInvitationRemoved:
                w.WriteBytesField(1, ReceivedInvitation?.Encode());
                w.WriteUInt32(3, Reason);
                if (SubscriberAccountId is not null) w.WriteBytesField(5, SubscriberAccountId.Encode());
                break;
            case BgsFriendsNotificationKind.SentInvitationAdded:
                if (SubscriberAccountId is not null) w.WriteBytesField(1, SubscriberAccountId.Encode());
                w.WriteBytesField(2, SentInvitation?.Encode());
                break;
            case BgsFriendsNotificationKind.SentInvitationRemoved:
                if (SubscriberAccountId is not null) w.WriteBytesField(1, SubscriberAccountId.Encode());
                w.WriteFixed64(2, InvitationId);
                w.WriteUInt32(3, Reason);
                break;
        }

        return w.ToArray();
    }

    /// <summary>Decodes the body of FriendsListener method <paramref name="methodId"/>; null for a method this client does not know.</summary>
    public static BgsFriendsNotification? Decode(uint methodId, byte[] data)
    {
        if (methodId is < 1 or > 7)
        {
            return null;
        }

        var kind = (BgsFriendsNotificationKind)methodId;
        BgsFriendMessage? friend = null;
        BgsReceivedInvitation? received = null;
        BgsSentInvitation? sent = null;
        ulong? invitationId = null;
        uint? reason = null;
        EntityId? subscriber = null;
        var sentForm = kind is BgsFriendsNotificationKind.SentInvitationAdded or BgsFriendsNotificationKind.SentInvitationRemoved;

        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            if (sentForm)
            {
                switch (field)
                {
                    case 1: subscriber = EntityId.Decode(r.ReadLengthDelimited()); break;
                    case 2 when kind == BgsFriendsNotificationKind.SentInvitationAdded: sent = BgsSentInvitation.Decode(r.ReadLengthDelimited()); break;
                    case 2 when type == WireType.Fixed64: invitationId = r.ReadFixed64(); break;
                    case 3: reason = (uint)r.ReadVarint(); break;
                    default: r.Skip(type); break;
                }

                continue;
            }

            switch (field)
            {
                case 1 when kind is BgsFriendsNotificationKind.ReceivedInvitationAdded or BgsFriendsNotificationKind.ReceivedInvitationRemoved:
                    received = BgsReceivedInvitation.Decode(r.ReadLengthDelimited());
                    break;
                case 1: friend = BgsFriendMessage.Decode(r.ReadLengthDelimited()); break;
                case 3 when type == WireType.Varint: reason = (uint)r.ReadVarint(); break;
                case 5: subscriber = EntityId.Decode(r.ReadLengthDelimited()); break;
                default: r.Skip(type); break; // legacy field 2 = game_account_id, unused
            }
        }

        return new BgsFriendsNotification
        {
            Kind = kind,
            Friend = friend,
            ReceivedInvitation = received,
            SentInvitation = sent,
            InvitationId = invitationId ?? received?.Id,
            Reason = reason,
            SubscriberAccountId = subscriber,
        };
    }
}

/// <summary>Packed/unpacked repeated varint helpers ProtoWriter/ProtoReader leave out.</summary>
internal static class ProtoPacking
{
    public static void WritePackedVarints(ProtoWriter w, int field, IReadOnlyCollection<uint> values)
    {
        if (values.Count == 0)
        {
            return;
        }

        var inner = new ProtoWriter();
        foreach (var v in values) inner.WriteVarint(v);
        w.WriteBytesField(field, inner.ToArray());
    }

    /// <summary>Accepts both the packed (length-delimited) and the unpacked (one varint) encodings.</summary>
    public static void ReadVarints(ProtoReader r, WireType type, List<uint> into)
    {
        if (type == WireType.LengthDelimited)
        {
            var inner = new ProtoReader(r.ReadLengthDelimited());
            while (inner.HasMore) into.Add((uint)inner.ReadVarint());
        }
        else if (type == WireType.Varint)
        {
            into.Add((uint)r.ReadVarint());
        }
        else
        {
            r.Skip(type);
        }
    }

    /// <summary>Repeated fixed32, packed or not.</summary>
    public static void ReadFixed32s(ProtoReader r, WireType type, List<uint> into)
    {
        if (type == WireType.LengthDelimited)
        {
            var inner = new ProtoReader(r.ReadLengthDelimited());
            while (inner.HasMore) into.Add(inner.ReadFixed32());
        }
        else if (type == WireType.Fixed32)
        {
            into.Add(r.ReadFixed32());
        }
        else
        {
            r.Skip(type);
        }
    }
}
