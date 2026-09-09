using Invigoration.Core.Chat;

namespace Invigoration.Core.Tests;

public class ChatIconTests
{
    [Theory]
    [InlineData("3RAW", "war3")]
    [InlineData("PX3W", "war3")]
    [InlineData("PX2D", "d2exp")]
    [InlineData("VD2D", "diablo2")]
    [InlineData("LTRD", "diablo")]
    [InlineData("RHSD", "dshr")]
    [InlineData("RATS", "sc")]
    [InlineData("PXES", "scbw")]
    [InlineData("RHSS", "sware")]
    [InlineData("RTSJ", "jsc")]
    [InlineData("NB2W", "war2")]
    // Regression: StatStringParser already recognizes "TAHC" as a Chat bot's product code, but
    // GetProductIconKey had no matching case, so a user connecting with that product silently
    // got no icon at all (the "chat" icon asset/catalog entry existed but was unreachable).
    [InlineData("TAHC", "chat")]
    [InlineData("sc2", "sc2")]
    public void GetProductIconKey_KnownProduct_ReturnsExpectedKey(string statString, string expected)
    {
        Assert.Equal(expected, ChatIcon.GetProductIconKey(statString));
    }

    [Fact]
    public void GetProductIconKey_UnknownProduct_ReturnsEmpty()
    {
        Assert.Equal("", ChatIcon.GetProductIconKey("ZZZZ"));
    }
}
