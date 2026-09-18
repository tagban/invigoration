using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>
/// The Bot menu's three-way joins/leaves choice is a view over the two switches the Config window
/// already had — these pin down that mapping in both directions, and that the choice survives a
/// save without being stored as a field of its own.
/// </summary>
public class BotConfigJoinLeaveDisplayTests
{
    [Theory]
    [InlineData(false, false, JoinLeaveDisplay.ShowAll)]
    [InlineData(false, true, JoinLeaveDisplay.HideSpam)]
    [InlineData(true, false, JoinLeaveDisplay.HideAll)]
    [InlineData(true, true, JoinLeaveDisplay.HideAll)]
    public void ReadsTheExistingSwitches(bool suppress, bool hideSpam, JoinLeaveDisplay expected)
    {
        var config = new BotConfig { SuppressJoinLeaveNotifications = suppress, HideJoinLeaveSpamEnabled = hideSpam };

        Assert.Equal(expected, config.JoinLeaveDisplay);
    }

    [Fact]
    public void NewBotHidesSpamByDefault() => Assert.Equal(JoinLeaveDisplay.HideSpam, new BotConfig().JoinLeaveDisplay);

    [Theory]
    [InlineData(JoinLeaveDisplay.ShowAll, false, false)]
    [InlineData(JoinLeaveDisplay.HideSpam, false, true)]
    public void ChoosingShowAllOrHideSpamSetsBothSwitches(JoinLeaveDisplay choice, bool suppress, bool hideSpam)
    {
        var config = new BotConfig { SuppressJoinLeaveNotifications = true, HideJoinLeaveSpamEnabled = !hideSpam };

        config.JoinLeaveDisplay = choice;

        Assert.Equal(suppress, config.SuppressJoinLeaveNotifications);
        Assert.Equal(hideSpam, config.HideJoinLeaveSpamEnabled);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void HidingEverythingLeavesTheSpamSwitchAlone(bool hideSpamBefore)
    {
        var config = new BotConfig { HideJoinLeaveSpamEnabled = hideSpamBefore };

        config.JoinLeaveDisplay = JoinLeaveDisplay.HideAll;

        Assert.True(config.SuppressJoinLeaveNotifications);
        Assert.Equal(hideSpamBefore, config.HideJoinLeaveSpamEnabled);
    }

    [Theory]
    [InlineData(JoinLeaveDisplay.ShowAll)]
    [InlineData(JoinLeaveDisplay.HideSpam)]
    [InlineData(JoinLeaveDisplay.HideAll)]
    public void SurvivesASaveThroughTheUnderlyingSwitches(JoinLeaveDisplay choice)
    {
        var config = new BotConfig { JoinLeaveDisplay = choice };

        var copy = BotConfig.Clone(config);

        Assert.Equal(choice, copy.JoinLeaveDisplay);
    }

    [Fact]
    public void IsNotWrittenAsAFieldOfItsOwn()
    {
        var json = System.Text.Json.JsonSerializer.Serialize(new BotConfig { JoinLeaveDisplay = JoinLeaveDisplay.HideAll });

        Assert.DoesNotContain(nameof(BotConfig.JoinLeaveDisplay), json);
    }
}
