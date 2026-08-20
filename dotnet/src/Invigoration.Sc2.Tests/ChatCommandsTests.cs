using Invigoration.Sc2.Chat;
using Invigoration.Sc2.Native;

namespace Invigoration.Sc2.Tests;

/// <summary>
/// Every hex vector here is reproduced verbatim from ncarrillo/superiority's
/// core/src/native/protocol.rs unit tests (relayed via research agent, then
/// cross-checked bit-for-bit against the toon_select checksum derivation by
/// hand before being trusted).
/// </summary>
public class ChatCommandsTests
{
    [Theory]
    [InlineData(0, "4205")]
    [InlineData(6, "4235")]
    public void ChatLeave_MatchesGoldenVector(byte channelIndex, string expectedHex)
    {
        var record = ChatCommands.ChatLeave(channelIndex);

        Assert.Equal(expectedHex, Convert.ToHexString(record).ToLowerInvariant());
    }

    [Fact]
    public void ChatLeave_OutOfRangeIndex_Throws()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ChatCommands.ChatLeave(7));
    }

    [Theory]
    [InlineData(0, true, "4505")]
    [InlineData(0, false, "4605")]
    [InlineData(6, true, "4535")]
    [InlineData(6, false, "4635")]
    public void ChatInviteAnswer_MatchesGoldenVector(byte channelIndex, bool accept, string expectedHex)
    {
        var record = ChatCommands.ChatInviteAnswer(channelIndex, accept);

        Assert.Equal(expectedHex, Convert.ToHexString(record).ToLowerInvariant());
    }

    [Fact]
    public void ChatInviteAnswer_OutOfRangeIndex_Throws()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ChatCommands.ChatInviteAnswer(7, true));
    }

    [Fact]
    public void ChatJoinPrivate_MatchesGoldenVector()
    {
        var record = ChatCommands.ChatJoinPrivate("Custom Room", 0x10203040);

        Assert.Equal("40050b437573746f6d20526f6f6d10203040", Convert.ToHexString(record).ToLowerInvariant());
    }

    [Fact]
    public void ChatJoinPrivate_EmptyName_Throws()
    {
        Assert.Throws<ArgumentException>(() => ChatCommands.ChatJoinPrivate("", 0));
    }

    [Fact]
    public void ChatJoinPrivate_TooManyCharacters_Throws()
    {
        Assert.Throws<ArgumentException>(() => ChatCommands.ChatJoinPrivate(new string('x', 32), 0));
    }

    [Fact]
    public void ChatJoinPrivate_31EmojiCharacters_Succeeds()
    {
        var name = string.Concat(Enumerable.Repeat("\U0001F6F0", 31)); // satellite emoji, 4 UTF-8 bytes each = 124 bytes

        var record = ChatCommands.ChatJoinPrivate(name, 0);

        Assert.NotEmpty(record);
    }

    [Fact]
    public void ToonSelect_MatchesGoldenVectorIncludingChecksum()
    {
        var record = ChatCommands.ToonSelect("hotshot#994", 1);

        Assert.Equal("c51701686f7473686f74233939348d0000000001", Convert.ToHexString(record).ToLowerInvariant());
    }

    [Fact]
    public void ChatWhisper_Presence_MatchesGoldenVector()
    {
        var record = ChatCommands.ChatWhisper(new WhisperTarget.Presence(0x02b75a16), ".");

        Assert.Equal("53050add6816012e", Convert.ToHexString(record).ToLowerInvariant());
    }
}
