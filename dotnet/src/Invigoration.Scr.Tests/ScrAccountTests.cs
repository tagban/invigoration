using Invigoration.Sc2.Protobuf;

namespace Invigoration.Scr.Tests;

/// <summary>Built from the layouts seen in a retail-client capture: GetToons' reply and GatewayUpdate.</summary>
public class ScrAccountTests
{
    private static readonly IReadOnlyList<ScrToon> Account =
    [
        new(1, "Invigoration2", ScrGateways.UsWest),
        new(2, "Tagban", ScrGateways.UsWest),
        new(3, "Tagban", ScrGateways.UsEast),
        new(4, "BNU-Master", ScrGateways.UsEast),
    ];

    [Fact]
    public void GetToons_DecodesEveryCharacterWithItsGateway()
    {
        var body = new ProtoWriter();
        foreach (var toon in Account)
        {
            var entry = new ProtoWriter();
            entry.WriteUInt32(1, toon.Id);
            entry.WriteString(2, toon.Name);
            entry.WriteUInt32(3, toon.Gateway);
            body.WriteBytesField(1, entry.ToArray());
        }

        Assert.Equal(Account, ScrToons.Decode(body.ToArray()));
    }

    [Theory]
    [InlineData(ScrGateways.UsEast, "", 3u)]
    [InlineData(ScrGateways.UsEast, "bnu-master", 4u)]
    [InlineData(ScrGateways.UsWest, "Tagban", 2u)]
    [InlineData(ScrGateways.UsWest, null, 1u)]
    public void Choose_PicksByGatewayThenName(uint gateway, string? name, uint expectedId) =>
        Assert.Equal(expectedId, ScrToons.Choose(Account, gateway, name).Id);

    [Fact]
    public void Choose_ExplainsAGatewayWithNoCharacter()
    {
        var error = Assert.Throws<ScrCharacterException>(() => ScrToons.Choose(Account, ScrGateways.Europe, null));

        Assert.Contains("no character on Europe", error.Message);
        Assert.Contains("Tagban (U.S. East)", error.Message);
    }

    [Fact]
    public void GatewayUpdate_DecodesIdAndName()
    {
        var gateway = new ProtoWriter();
        gateway.WriteUInt32(1, 20);
        gateway.WriteString(2, "Europe");
        gateway.WriteString(3, "eu");
        gateway.WriteUInt32(5, 1);
        var body = new ProtoWriter();
        body.WriteBytesField(1, gateway.ToArray());
        body.WriteUInt32(2, 0);

        Assert.Equal(new ScrGateway(20, "Europe"), ScrGateways.DecodeUpdate(body.ToArray()));
    }
}

public class ScrFriendsTests
{
    [Fact]
    public void FriendUpdated_KeepsTagProgramAndStateButNotTheRealName()
    {
        var friend = new ProtoWriter();
        friend.WriteUInt64(1, 374291212);
        friend.WriteString(2, "islanti#11308");
        friend.WriteString(3, "Real Name");
        friend.WriteString(4, "S1");
        friend.WriteUInt32(5, 1);
        friend.WriteUInt32(6, 0);
        friend.WriteUInt32(7, 1);
        var body = new ProtoWriter();
        body.WriteBytesField(1, friend.ToArray());
        body.WriteUInt32(2, 0);

        var (decoded, removed) = ScrFriends.DecodeUpdate(body.ToArray())!.Value;

        Assert.Equal(new ScrFriend(374291212, "islanti#11308", "S1", true, false, true), decoded);
        Assert.False(removed);
        Assert.DoesNotContain("Real Name", decoded.ToString());
    }
}
