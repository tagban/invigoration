using Invigoration.Core.Commands;

namespace Invigoration.Core.Tests;

public class CommandCatalogTypoSuggestionTests
{
    [Theory]
    [InlineData("idel", "idle")]
    [InlineData("idl", "idle")]
    [InlineData("hepl", "help")]
    [InlineData("uptme", "uptime")]
    public void SuggestClosestAlias_ForACloseTypo_SuggestsTheRealCommand(string typo, string expected)
    {
        Assert.Equal(expected, CommandCatalog.SuggestClosestAlias(typo));
    }

    [Theory]
    [InlineData("whois")]
    [InlineData("f")]
    [InlineData("w")]
    [InlineData("friends")]
    public void SuggestClosestAlias_ForARealBncsCommand_SuggestsNothing(string realServerCommand)
    {
        Assert.Null(CommandCatalog.SuggestClosestAlias(realServerCommand));
    }

    [Fact]
    public void SuggestClosestAlias_ForAnExactMatch_SuggestsNothing()
    {
        Assert.Null(CommandCatalog.SuggestClosestAlias("idle"));
    }
}
