using System.Text;
using Invigoration.Core.Networking;
using Invigoration.Core.Protocol;

namespace Invigoration.Core.Chat;

/// <summary>
/// Parses the classic BNCS friends-list packets, field layouts per
/// bnetdocs.org: SID_FRIENDSLIST (0x65, S&gt;C — the full list),
/// SID_FRIENDSUPDATE (0x66, S&gt;C — one entry's status changed),
/// SID_FRIENDSADD (0x67), SID_FRIENDSREMOVE (0x68), SID_FRIENDSPOSITION (0x69).
/// Adding/removing/reordering friends themselves is done via plain
/// "/f add|remove|promote|demote &lt;name&gt;" chat text handled server-side (the short
/// "/f" alias, not "/friend" — the full word isn't recognized on real Battle.net, confirmed live:
/// "/friend add X" got "That is not a valid command" from useast.battle.net) —
/// there's no dedicated outbound packet for it, unlike these five replies.
/// </summary>
public static class FriendsListParser
{
    public static IReadOnlyList<FriendEntry> ParseFriendsList(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        var count = reader.ReadByte();
        var friends = new List<FriendEntry>(count);
        for (var i = 0; i < count; i++)
        {
            friends.Add(ReadEntry(reader));
        }

        return friends;
    }

    public static (byte EntryNumber, FriendStatusUpdate Update) ParseFriendsUpdate(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        var entryNumber = reader.ReadByte();
        return (entryNumber, ReadStatusUpdate(reader));
    }

    public static FriendEntry ParseFriendsAdd(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        return ReadEntry(reader);
    }

    public static byte ParseFriendsRemove(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        return reader.ReadByte();
    }

    /// <summary>Returns (old position, new position); the caller should shift every entry between them, per bnetdocs' description of this packet.</summary>
    public static (byte OldEntry, byte NewEntry) ParseFriendsPosition(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        return (reader.ReadByte(), reader.ReadByte());
    }

    private static FriendEntry ReadEntry(PacketReader reader)
    {
        var account = reader.ReadNTString();
        var update = ReadStatusUpdate(reader);
        return new FriendEntry(account, update.Status, update.Location, update.ProductCode, update.LocationName);
    }

    private static FriendStatusUpdate ReadStatusUpdate(PacketReader reader)
    {
        var status = (FriendStatus)reader.ReadByte();
        var location = (FriendLocation)reader.ReadByte();
        var productCode = Encoding.Latin1.GetString(reader.ReadRaw(4));
        var locationName = reader.ReadNTString();
        return new FriendStatusUpdate(status, location, productCode, locationName);
    }
}
