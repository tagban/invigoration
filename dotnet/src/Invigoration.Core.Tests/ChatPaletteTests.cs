using Invigoration.Core.Chat;

namespace Invigoration.Core.Tests;

public class ChatPaletteTests
{
    private static readonly ChatPalette Palette = ChatPalette.Invigoration;

    // Real classic Battle.net channel-list username coloring, per explicit user description:
    // Blizzard Reps blue, Battle.net Admins green, same-game-as-you white, everyone else yellow.
    [Fact]
    public void GetChannelListNameColor_Blizzard_ReturnsBlue()
    {
        Assert.Equal(Palette.Blue, Palette.GetChannelListNameColor((uint)UserFlags.Blizzard, isSameGame: false));
    }

    [Fact]
    public void GetChannelListNameColor_Admin_ReturnsGreen()
    {
        Assert.Equal(Palette.Green, Palette.GetChannelListNameColor((uint)UserFlags.Admin, isSameGame: false));
    }

    [Fact]
    public void GetChannelListNameColor_SameGame_ReturnsWhite()
    {
        Assert.Equal(Palette.White, Palette.GetChannelListNameColor((uint)UserFlags.None, isSameGame: true));
    }

    [Fact]
    public void GetChannelListNameColor_DifferentGameNoFlags_ReturnsYellow()
    {
        Assert.Equal(Palette.Yellow, Palette.GetChannelListNameColor((uint)UserFlags.None, isSameGame: false));
    }

    // Regression: unlike GetUserNameColor (used for chat text), Operator/Speaker/Special guest
    // don't get their own channel-list color — real Battle.net only distinguished Blizzard/Admin/
    // same-game/everyone-else there, since rank badges already convey the rest visually.
    [Theory]
    [InlineData(UserFlags.Operator)]
    [InlineData(UserFlags.Speaker)]
    [InlineData(UserFlags.Special)]
    [InlineData(UserFlags.Squelched)]
    public void GetChannelListNameColor_OtherFlagsAlone_FallBackToGameCheck(UserFlags flag)
    {
        Assert.Equal(Palette.Yellow, Palette.GetChannelListNameColor((uint)flag, isSameGame: false));
        Assert.Equal(Palette.White, Palette.GetChannelListNameColor((uint)flag, isSameGame: true));
    }

    [Fact]
    public void GetChannelListNameColor_BlizzardBeatsAdminAndSameGame()
    {
        var flags = (uint)(UserFlags.Blizzard | UserFlags.Admin);
        Assert.Equal(Palette.Blue, Palette.GetChannelListNameColor(flags, isSameGame: true));
    }
}
