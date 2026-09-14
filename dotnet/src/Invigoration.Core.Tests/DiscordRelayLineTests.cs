using Invigoration.Core.Discord;

namespace Invigoration.Core.Tests;

public class DiscordRelayLineTests
{
    [Fact]
    public void Format_ThenTryParse_RoundTrips()
    {
        var line = DiscordRelayLine.Format("dave", "anyone up for a baal run? 10:30 works");

        Assert.Equal("[Discord] dave: anyone up for a baal run? 10:30 works", line);
        Assert.True(DiscordRelayLine.TryParse(line, out var user, out var message));
        Assert.Equal("dave", user);
        Assert.Equal("anyone up for a baal run? 10:30 works", message);
    }

    [Theory]
    [InlineData("hello")]
    [InlineData("[discord] dave: hi")]          // prefix is exact
    [InlineData("[Discord] dave")]              // no separator
    [InlineData("[Discord] : hi")]              // no name
    [InlineData("[Discord]    : hi")]           // blank name
    [InlineData("xx [Discord] dave: hi")]       // not at the start
    public void TryParse_RejectsAnythingElse(string text)
    {
        Assert.False(DiscordRelayLine.TryParse(text, out _, out _));
    }

    [Fact]
    public void TryParse_RejectsImplausiblyLongNames()
    {
        Assert.False(DiscordRelayLine.TryParse("[Discord] " + new string('a', 60) + ": hi", out _, out _));
    }

    [Fact]
    public void TryParse_AllowsAnEmptyMessage()
    {
        Assert.True(DiscordRelayLine.TryParse("[Discord] dave: ", out var user, out var message));
        Assert.Equal("dave", user);
        Assert.Equal("", message);
    }

    [Theory]
    [InlineData("[Discord] dave", "dave")]
    [InlineData("Tagban", null)]
    [InlineData("[Discord] ", null)]
    public void DiscordUserFromSpeaker_ReadsTheSpeakerForm(string speaker, string? expected)
    {
        Assert.Equal(expected, DiscordRelayLine.DiscordUserFromSpeaker(speaker));
    }
}
