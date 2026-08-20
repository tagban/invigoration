using Invigoration.Core.Chat;
using Invigoration.Core.Protocol;

namespace Invigoration.Core.Tests;

public class FriendsListParserTests
{
    [Fact]
    public void ParseFriendsList_ReadsEachEntryInOrder()
    {
        var frame = new PacketWriter()
            .WriteByte(2) // count
            .WriteNTString("alice")
            .WriteByte((byte)(FriendStatus.Mutual | FriendStatus.Away))
            .WriteByte((byte)FriendLocation.InChat)
            .WriteAscii("VD2D")
            .WriteNTString("Some Channel")
            .WriteNTString("bob")
            .WriteByte((byte)FriendStatus.None)
            .WriteByte((byte)FriendLocation.Offline)
            .WriteAscii("PX2D")
            .WriteNTString("")
            .ToBncsPacket(BncsPacketId.SID_FRIENDSLIST);

        var friends = FriendsListParser.ParseFriendsList(frame);

        Assert.Equal(2, friends.Count);
        Assert.Equal(new FriendEntry("alice", FriendStatus.Mutual | FriendStatus.Away, FriendLocation.InChat, "VD2D", "Some Channel"), friends[0]);
        Assert.Equal(new FriendEntry("bob", FriendStatus.None, FriendLocation.Offline, "PX2D", ""), friends[1]);
    }

    [Fact]
    public void ParseFriendsUpdate_ReadsEntryNumberAndStatusWithNoAccountField()
    {
        var frame = new PacketWriter()
            .WriteByte(3) // entry number
            .WriteByte((byte)FriendStatus.Mutual)
            .WriteByte((byte)FriendLocation.PublicGame)
            .WriteAscii("RATS")
            .WriteNTString("some game")
            .ToBncsPacket(BncsPacketId.SID_FRIENDSUPDATE);

        var (entryNumber, update) = FriendsListParser.ParseFriendsUpdate(frame);

        Assert.Equal(3, entryNumber);
        Assert.Equal(FriendStatus.Mutual, update.Status);
        Assert.Equal(FriendLocation.PublicGame, update.Location);
        Assert.Equal("RATS", update.ProductCode);
        Assert.Equal("some game", update.LocationName);
    }

    [Fact]
    public void ParseFriendsAdd_ReadsFullEntry()
    {
        var frame = new PacketWriter()
            .WriteNTString("newfriend")
            .WriteByte((byte)FriendStatus.Away)
            .WriteByte((byte)FriendLocation.NotInChat)
            .WriteAscii("PXES")
            .WriteNTString("")
            .ToBncsPacket(BncsPacketId.SID_FRIENDSADD);

        var entry = FriendsListParser.ParseFriendsAdd(frame);

        Assert.Equal(new FriendEntry("newfriend", FriendStatus.Away, FriendLocation.NotInChat, "PXES", ""), entry);
    }

    [Fact]
    public void ParseFriendsRemove_ReadsEntryNumber()
    {
        var frame = new PacketWriter().WriteByte(1).ToBncsPacket(BncsPacketId.SID_FRIENDSREMOVE);

        var entryNumber = FriendsListParser.ParseFriendsRemove(frame);

        Assert.Equal(1, entryNumber);
    }

    [Fact]
    public void ParseFriendsPosition_ReadsOldAndNewEntry()
    {
        var frame = new PacketWriter().WriteByte(4).WriteByte(0).ToBncsPacket(BncsPacketId.SID_FRIENDSPOSITION);

        var (oldEntry, newEntry) = FriendsListParser.ParseFriendsPosition(frame);

        Assert.Equal(4, oldEntry);
        Assert.Equal(0, newEntry);
    }
}
