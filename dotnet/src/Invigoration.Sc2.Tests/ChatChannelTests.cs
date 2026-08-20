using Invigoration.Sc2.Chat;

namespace Invigoration.Sc2.Tests;

public class ChatChannelTests
{
    [Fact]
    public void Title_DefaultPublicChannel_IsGeneral()
    {
        Assert.Equal("General", ChatChannel.DefaultPublic().Title());
    }

    [Fact]
    public void Title_OtherPublicChannel_ShowsId()
    {
        Assert.Equal("Public 42", new ChatChannel.Public(42).Title());
    }

    [Fact]
    public void Title_Private_ShowsName()
    {
        Assert.Equal("my-channel", new ChatChannel.Private("my-channel").Title());
    }

    [Fact]
    public void Title_Club_ShowsGroupId()
    {
        Assert.Equal("Group 7", new ChatChannel.Club(7).Title());
    }

    [Fact]
    public void Title_Party_IsParty()
    {
        Assert.Equal("Party", new ChatChannel.Party().Title());
    }
}

public class ChatUserTests
{
    [Fact]
    public void VisibleName_StripsBattleTagDiscriminator()
    {
        var user = new ChatUser(1, null, "Tagban#1234", null, PresenceState.Online);

        Assert.Equal("Tagban", user.VisibleName());
    }

    [Fact]
    public void VisibleName_PrefixesClanTag()
    {
        var user = new ChatUser(1, null, "Tagban#1234", "BNU", PresenceState.Online);

        Assert.Equal("<BNU>Tagban", user.VisibleName());
    }

    [Fact]
    public void VisibleName_MissingName_FallsBackToHandle()
    {
        var user = new ChatUser(99, null, null, null, PresenceState.Unknown);

        Assert.Equal("Player 99", user.VisibleName());
    }
}
