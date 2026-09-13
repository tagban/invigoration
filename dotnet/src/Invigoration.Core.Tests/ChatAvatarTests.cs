using Invigoration.Core.Chat;
using Invigoration.Core.StatString;

namespace Invigoration.Core.Tests;

public class ChatAvatarTests
{
    private static string D2(string product, int classRaw, int flags, int actByte = 0x80)
    {
        var p = Enumerable.Repeat((char)0xFF, 33).ToArray();
        p[0] = (char)0x84;
        p[1] = (char)0x80;
        p[13] = (char)classRaw;
        p[25] = (char)50;
        p[26] = (char)flags;
        p[27] = (char)actByte;
        return product + "USEast,Kilua," + new string(p);
    }

    [Theory]
    [InlineData("RATS 0 0 7 0 0 0 0 0 RATS", ChatAvatar.StarCraft)]
    [InlineData("RTSJ 0 0 0 0 0 0 0 0 RTSJ", ChatAvatar.StarCraft)]
    [InlineData("PXES 0 0 0 0 0 0 0 0 PXES", ChatAvatar.BroodWar)]
    [InlineData("NB2W 0 0 0 0 0 0 0 0 NB2W", ChatAvatar.WarcraftII)]
    [InlineData("LTRD 5 1 0 30 20 20 30 100 0", ChatAvatar.Diablo)]
    [InlineData("TAHC", ChatAvatar.ChatClient)]
    [InlineData("3RAW 5H3W 20", ChatAvatar.Unknown)]
    [InlineData("", ChatAvatar.Unknown)]
    public void OtherClients_GetTheirProductAvatar(string statString, string expected)
    {
        Assert.Equal(expected, ChatAvatar.For(0, statString));
    }

    [Fact]
    public void ALiveD2Character_IsDrawnAsThemselves()
    {
        Assert.Null(ChatAvatar.For(0, D2("PX2D", 4, 0xA0)));
    }

    // Per The Arreat Summit: "Open Characters clients also appear as this Avatar."
    [Theory]
    [InlineData("PX2D")]
    [InlineData("VD2D")]
    public void AnOpenD2Character_IsTheUnknownAvatar(string statString)
    {
        Assert.Equal(ChatAvatar.Unknown, ChatAvatar.For(0, statString));
    }

    [Fact]
    public void ADeadHardcoreCharacter_IsTheGhost()
    {
        Assert.Equal(ChatAvatar.DeadHardcore, ChatAvatar.For(0, D2("PX2D", 4, 0xA0 | 0x04 | 0x08)));
        // Softcore death isn't permanent — still themselves.
        Assert.Null(ChatAvatar.For(0, D2("PX2D", 4, 0xA0 | 0x08)));
    }

    // Rank outranks the client, so a D2 player running the channel shows as the Moderator.
    [Theory]
    [InlineData(UserFlags.Blizzard | UserFlags.Operator, ChatAvatar.BlizzardRep)]
    [InlineData(UserFlags.Admin | UserFlags.Operator, ChatAvatar.Sysop)]
    [InlineData(UserFlags.Operator, ChatAvatar.Moderator)]
    [InlineData(UserFlags.Speaker, ChatAvatar.Speaker)]
    [InlineData((UserFlags)0x8000, ChatAvatar.Referee)]
    public void Rank_WinsOverTheClient(UserFlags flags, string expected)
    {
        Assert.Equal(expected, ChatAvatar.For((uint)flags, D2("PX2D", 4, 0xA0)));
        Assert.Equal(expected, ChatAvatar.For((uint)flags, "RATS 0 0 0 0 0 0 0 0 RATS"));
    }

    [Fact]
    public void SquelchedAndSpecialGuest_DontChangeTheAvatar()
    {
        Assert.Equal(ChatAvatar.StarCraft, ChatAvatar.For((uint)(UserFlags.Squelched | UserFlags.Special), "RATS 0 0 0 0 0 0 0 0 RATS"));
    }

    [Theory]
    [InlineData(0x80, 0xA0, 1, "Kilua")]              // LoD amazon, Normal in progress
    [InlineData(0x8A, 0xA0, 1, "Slayer Kilua")]
    [InlineData(0x9E, 0xA0, 1, "Matriarch Kilua")]
    [InlineData(0x9E, 0xA0, 4, "Patriarch Kilua")]
    [InlineData(0x9E, 0xA4, 1, "Guardian Kilua")]
    [InlineData(0x88, 0x80, 2, "Dame Kilua")]         // classic sorceress
    [InlineData(0x98, 0x84, 5, "King Kilua")]         // classic hardcore barbarian, all done
    public void TitledName_MatchesTheLobbyLabel(int actByte, int flags, int classRaw, string expected)
    {
        Assert.True(D2Character.TryParse(D2(flags >= 0xA0 ? "PX2D" : "VD2D", classRaw, flags, actByte), out var c));
        Assert.Equal(expected, c.TitledName);
    }
}
