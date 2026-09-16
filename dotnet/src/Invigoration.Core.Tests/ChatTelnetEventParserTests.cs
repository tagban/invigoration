using Invigoration.Core.Chat;

namespace Invigoration.Core.Tests;

/// <summary>Cases 1-3 are lifted verbatim from a live capture the user supplied; the rest exercise the parser's own edge-case handling.</summary>
public class ChatTelnetEventParserTests
{
    [Fact]
    public void TryParse_UserEvent_ParsesUsernameFlagsAndBracketedTag()
    {
        var result = ChatTelnetEventParser.TryParse("1001 USER Jailout2000 0010 [CHAT]");

        Assert.NotNull(result);
        Assert.Equal(ChatEventType.ShowUser, result!.Type);
        Assert.Equal("Jailout2000", result.Username);
        Assert.Equal(0x0010u, result.Flags);
        Assert.Equal("[CHAT]", result.Text);
    }

    [Fact]
    public void TryParse_ChannelEvent_UnquotesNameWithNoUsernameOrFlags()
    {
        var result = ChatTelnetEventParser.TryParse("1007 CHANNEL \"Public Chat 1\"");

        Assert.NotNull(result);
        Assert.Equal(ChatEventType.Channel, result!.Type);
        Assert.Equal("", result.Username);
        Assert.Equal(0u, result.Flags);
        Assert.Equal("Public Chat 1", result.Text);
    }

    [Fact]
    public void TryParse_TalkEvent_UnquotesMessageWithEmbeddedApostrophe()
    {
        var result = ChatTelnetEventParser.TryParse("1005 TALK Jailout2000 0010 \"It's just me in this channel.\"");

        Assert.NotNull(result);
        Assert.Equal(ChatEventType.Talk, result!.Type);
        Assert.Equal("Jailout2000", result.Username);
        Assert.Equal(0x0010u, result.Flags);
        Assert.Equal("It's just me in this channel.", result.Text);
    }

    [Fact]
    public void TryParse_JoinEvent_MapsToSamePatternAsUser()
    {
        var result = ChatTelnetEventParser.TryParse("1002 JOIN SomeoneElse 0000 [CHAT]");

        Assert.NotNull(result);
        Assert.Equal(ChatEventType.Join, result!.Type);
        Assert.Equal("SomeoneElse", result.Username);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("not a numbered line")]
    [InlineData("9999 UNKNOWN whoever 0000 text")]
    public void TryParse_UnrecognizedOrBlankLines_ReturnsNull(string line)
    {
        Assert.Null(ChatTelnetEventParser.TryParse(line));
    }

    [Fact]
    public void TryParse_NameConfirmationLine_ReturnsNull()
    {
        // 2010 NAME is handled separately by BotEngine.Chat.cs's login handshake, not as a ChatEvent.
        Assert.Null(ChatTelnetEventParser.TryParse("2010 NAME Jailout2000"));
    }

    [Fact]
    public void TryParse_HexFlagsWithLetters_ParsesCorrectly()
    {
        var result = ChatTelnetEventParser.TryParse("1009 USERFLAGS SomeUser 001F [CHAT]");

        Assert.NotNull(result);
        Assert.Equal(0x001Fu, result!.Flags);
    }
}
