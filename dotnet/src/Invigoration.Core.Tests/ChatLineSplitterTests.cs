using Invigoration.Core.Chat;

namespace Invigoration.Core.Tests;

public class ChatLineSplitterTests
{
    [Fact]
    public void ShortText_IsSentAsIs()
    {
        Assert.Equal(["hello there"], ChatLineSplitter.Split("hello there"));
    }

    [Fact]
    public void EmptyText_SendsNothing()
    {
        Assert.Empty(ChatLineSplitter.Split(""));
    }

    [Fact]
    public void ExactlyTheLimit_IsOneLine()
    {
        var text = new string('a', ChatLineSplitter.MaxLineLength);
        Assert.Equal([text], ChatLineSplitter.Split(text));
    }

    [Fact]
    public void LongChat_WrapsAtWordBoundaries_WithinTheLimit()
    {
        var text = string.Join(' ', Enumerable.Range(0, 60).Select(i => $"word{i:00}"));

        var lines = ChatLineSplitter.Split(text);

        Assert.True(lines.Count > 1);
        Assert.All(lines, l => Assert.InRange(l.Length, 1, ChatLineSplitter.MaxLineLength));
        Assert.All(lines, l => Assert.False(l.StartsWith(' ') || l.EndsWith(' ')));
        Assert.Equal(text, string.Join(' ', lines));
    }

    [Theory]
    [InlineData("/w Tagban ")]
    [InlineData("/whisper Tagban ")]
    [InlineData("/me ")]
    [InlineData("/emote ")]
    public void LongWhisperOrEmote_RepeatsItsPrefixOnEveryLine(string prefix)
    {
        var body = string.Join(' ', Enumerable.Range(0, 60).Select(i => $"word{i:00}"));

        var lines = ChatLineSplitter.Split(prefix + body);

        Assert.True(lines.Count > 1);
        Assert.All(lines, l => Assert.StartsWith(prefix, l));
        Assert.All(lines, l => Assert.InRange(l.Length, prefix.Length + 1, ChatLineSplitter.MaxLineLength));
        Assert.Equal(body, string.Join(' ', lines.Select(l => l[prefix.Length..])));
    }

    [Fact]
    public void AWordLongerThanALine_IsCut()
    {
        var text = new string('x', 500);

        var lines = ChatLineSplitter.Split(text);

        Assert.Equal(3, lines.Count);
        Assert.All(lines, l => Assert.InRange(l.Length, 1, ChatLineSplitter.MaxLineLength));
        Assert.Equal(text, string.Concat(lines));
    }

    // Any other command can't be split into several meaningful commands, so it's cut to fit.
    [Fact]
    public void OtherLongCommand_IsTruncated()
    {
        var text = "/ban Tagban " + new string('r', 300);

        var lines = ChatLineSplitter.Split(text);

        Assert.Equal([text[..ChatLineSplitter.MaxLineLength]], lines);
    }
}
