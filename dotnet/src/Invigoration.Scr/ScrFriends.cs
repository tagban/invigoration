using Invigoration.Sc2.Protobuf;

namespace Invigoration.Scr;

/// <summary>A Battle.net friend as AuroraFriends.FriendUpdated describes them.</summary>
/// <param name="RealName">Their real name, when they share it (Real ID friends); empty otherwise.</param>
/// <param name="Detail">Fields 8 and 9 as text when they hold any; taken to be what the game says they're doing. Not confirmed.</param>
/// <param name="Program">The game or app they're in ("S1", "S2", "BSAp" for the Battle.net app...), empty when offline.</param>
/// <param name="Away">Field 6, set for one friend in a live session; taken to be away. Not confirmed.</param>
/// <param name="Busy">Field 7, likewise; taken to be busy. Not confirmed.</param>
public sealed record ScrFriend(ulong AccountId, string BattleTag, string Program, bool Online, bool Away, bool Busy, string RealName = "", string Detail = "");

/// <summary>A friend request (AuroraFriends.InvitationUpdated), answered by its <see cref="Id"/>.</summary>
public sealed record ScrInvitation(ulong Id, string BattleTag);

public static class ScrFriends
{
    public const uint Service = 0xAA4E1E00;
    public const uint FriendUpdatedMethod = 0xEC7E2FD1;

    // From the retail client's library, where each is registered next to its name.
    public const uint SendInvitationMethod = 0xF7C61139;
    public const uint RemoveFriendMethod = 0xCA10E52C;
    public const uint AcceptInvitationMethod = 0xF7E2F4AB;
    public const uint DeclineInvitationMethod = 0xE10265CB;
    public const uint InvitationUpdatedMethod = 0x4A8E2E5E;

    /// <summary>SendInvitation: {1: BattleTag}. The reply carries one string, meaning unknown.</summary>
    public static byte[] SendInvitationRequest(string battleTag)
    {
        var request = new ProtoWriter();
        request.WriteString(1, battleTag);
        return request.ToArray();
    }

    /// <summary>RemoveFriend: {1: account ID}, a plain varint (unlike a whisper's fixed32).</summary>
    public static byte[] RemoveFriendRequest(uint accountId)
    {
        var request = new ProtoWriter();
        request.WriteUInt32(1, accountId);
        return request.ToArray();
    }

    /// <summary>AcceptInvitation and DeclineInvitation: {1: invitation ID}.</summary>
    public static byte[] AnswerInvitationRequest(ulong invitationId)
    {
        var request = new ProtoWriter();
        request.WriteUInt64(1, invitationId);
        return request.ToArray();
    }

    /// <summary>InvitationUpdated: {1: {1: invitation ID, 2: BattleTag}, 2: removed}.</summary>
    public static (ScrInvitation Invitation, bool Removed)? DecodeInvitation(byte[] body)
    {
        ScrInvitation? invitation = null;
        var removed = false;
        var outer = new ProtoReader(body);
        while (outer.HasMore)
        {
            var (field, type) = outer.ReadTag();
            if (field == 1 && type == WireType.LengthDelimited)
            {
                ulong id = 0;
                var tag = "";
                var r = new ProtoReader(outer.ReadLengthDelimited());
                while (r.HasMore)
                {
                    var (inner, innerType) = r.ReadTag();
                    switch (inner)
                    {
                        case 1 when innerType == WireType.Varint:
                            id = r.ReadVarint();
                            break;
                        case 2 when innerType == WireType.LengthDelimited:
                            tag = r.ReadString();
                            break;
                        default:
                            r.Skip(innerType);
                            break;
                    }
                }

                invitation = new ScrInvitation(id, tag);
            }
            else if (field == 2 && type == WireType.Varint)
            {
                removed = outer.ReadVarint() != 0;
            }
            else
            {
                outer.Skip(type);
            }
        }

        return invitation is { Id: > 0 } ? (invitation, removed) : null;
    }

    /// <summary>
    /// FriendUpdated: {1: {1: account ID, 2: BattleTag, 3: real name, 4: program, 5: online,
    /// 6: away?, 7: busy?, 8/9: detail?}, 2: change}. Battle.net sends one per friend at sign-in and again as
    /// they change; a change of 1 is taken to be a removal (every one seen so far is 0).
    /// </summary>
    public static (ScrFriend Friend, bool Removed)? DecodeUpdate(byte[] body)
    {
        ScrFriend? friend = null;
        ulong change = 0;
        var outer = new ProtoReader(body);
        while (outer.HasMore)
        {
            var (field, type) = outer.ReadTag();
            if (field == 1 && type == WireType.LengthDelimited)
            {
                friend = DecodeFriend(outer.ReadLengthDelimited());
            }
            else if (field == 2 && type == WireType.Varint)
            {
                change = outer.ReadVarint();
            }
            else
            {
                outer.Skip(type);
            }
        }

        return friend is null || friend.BattleTag.Length == 0 ? null : (friend, change == 1);
    }

    private static ScrFriend DecodeFriend(byte[] body)
    {
        ulong id = 0;
        string tag = "", program = "", realName = "";
        var detail = new List<string>();
        bool online = false, away = false, busy = false;
        var r = new ProtoReader(body);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1 when type == WireType.Varint:
                    id = r.ReadVarint();
                    break;
                case 2 when type == WireType.LengthDelimited:
                    tag = r.ReadString();
                    break;
                case 3 when type == WireType.LengthDelimited:
                    realName = r.ReadString();
                    break;
                case 4 when type == WireType.LengthDelimited:
                    program = r.ReadString();
                    break;
                case 8 or 9 when type == WireType.LengthDelimited:
                    if (AsText(r.ReadLengthDelimited()) is { } text)
                    {
                        detail.Add(text);
                    }

                    break;
                case 5 when type == WireType.Varint:
                    online = r.ReadVarint() != 0;
                    break;
                case 6 when type == WireType.Varint:
                    away = r.ReadVarint() != 0;
                    break;
                case 7 when type == WireType.Varint:
                    busy = r.ReadVarint() != 0;
                    break;
                default:
                    r.Skip(type);
                    break;
            }
        }

        return new ScrFriend(id, tag, program, online, away, busy, realName.Trim(), string.Join(" ", detail));
    }

    private static string? AsText(byte[] bytes)
    {
        if (bytes.Length == 0)
        {
            return null;
        }

        try
        {
            var text = new System.Text.UTF8Encoding(false, true).GetString(bytes).Trim();
            return text.Length > 0 && !text.Any(char.IsControl) ? text : null;
        }
        catch (System.Text.DecoderFallbackException)
        {
            return null;
        }
    }
}
