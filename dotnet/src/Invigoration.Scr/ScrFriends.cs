using Invigoration.Sc2.Protobuf;

namespace Invigoration.Scr;

/// <summary>
/// A Battle.net friend as AuroraFriends.FriendUpdated describes them. The account's real name is
/// in the same message and deliberately not kept.
/// </summary>
/// <param name="Program">The game or app they're in ("S1", "S2", "BSAp" for the Battle.net app...), empty when offline.</param>
/// <param name="Away">Field 6, set for one friend in a live session; taken to be away. Not confirmed.</param>
/// <param name="Busy">Field 7, likewise; taken to be busy. Not confirmed.</param>
public sealed record ScrFriend(ulong AccountId, string BattleTag, string Program, bool Online, bool Away, bool Busy);

public static class ScrFriends
{
    public const uint Service = 0xAA4E1E00;
    public const uint FriendUpdatedMethod = 0xEC7E2FD1;

    /// <summary>
    /// FriendUpdated: {1: {1: account ID, 2: BattleTag, 3: real name, 4: program, 5: online,
    /// 6: away?, 7: busy?}, 2: change}. Battle.net sends one per friend at sign-in and again as
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
        string tag = "", program = "";
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
                case 4 when type == WireType.LengthDelimited:
                    program = r.ReadString();
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

        return new ScrFriend(id, tag, program, online, away, busy);
    }
}
