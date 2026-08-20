using Invigoration.Core.Config;
using Invigoration.Core.Trivia;

namespace Invigoration.Core.Tests;

public class TriviaGroupRegistryTests
{
    [Fact]
    public void GetSession_SameGroupName_ReturnsSameInstance()
    {
        var groupName = $"group-{Guid.NewGuid():N}";

        var first = TriviaGroupRegistry.GetSession(groupName);
        var second = TriviaGroupRegistry.GetSession(groupName);

        Assert.Same(first, second);
    }

    [Fact]
    public void GetSession_DifferentGroupNames_ReturnsDifferentInstances()
    {
        var first = TriviaGroupRegistry.GetSession($"group-{Guid.NewGuid():N}");
        var second = TriviaGroupRegistry.GetSession($"group-{Guid.NewGuid():N}");

        Assert.NotSame(first, second);
    }

    [Fact]
    public void GetSession_GroupNameIsCaseInsensitive()
    {
        var suffix = Guid.NewGuid().ToString("N");

        var lower = TriviaGroupRegistry.GetSession($"group-{suffix}");
        var upper = TriviaGroupRegistry.GetSession($"GROUP-{suffix}".ToUpperInvariant());

        Assert.Same(lower, upper);
    }

    [Fact]
    public async Task GetGroupPeers_FindsOtherEnginesInSameGroup_ExcludingSelf()
    {
        var groupName = $"group-{Guid.NewGuid():N}";
        await using var a = new BotEngine(new BotConfig { TriviaGroup = groupName });
        await using var b = new BotEngine(new BotConfig { TriviaGroup = groupName });
        await using var unrelated = new BotEngine(new BotConfig { TriviaGroup = $"other-{Guid.NewGuid():N}" });

        var peersOfA = TriviaGroupRegistry.GetGroupPeers(groupName, a);

        Assert.Single(peersOfA);
        Assert.Same(b, peersOfA[0]);
        Assert.DoesNotContain(unrelated, peersOfA);
    }

    [Fact]
    public async Task GetGroupPeers_UngroupedEngine_UsesCurrentConfigNotSnapshot()
    {
        var groupName = $"group-{Guid.NewGuid():N}";
        var config = new BotConfig { TriviaGroup = "" };
        await using var a = new BotEngine(config);
        await using var b = new BotEngine(new BotConfig { TriviaGroup = groupName });

        Assert.Empty(TriviaGroupRegistry.GetGroupPeers(groupName, b));

        // Simulate an operator editing the bot's config after the engine was constructed.
        config.TriviaGroup = groupName;

        var peersOfB = TriviaGroupRegistry.GetGroupPeers(groupName, b);
        Assert.Single(peersOfB);
        Assert.Same(a, peersOfB[0]);
    }

    [Fact]
    public async Task DisposeAsync_RemovesEngineFromGroupPeers()
    {
        var groupName = $"group-{Guid.NewGuid():N}";
        var b = new BotEngine(new BotConfig { TriviaGroup = groupName });
        await using var a = new BotEngine(new BotConfig { TriviaGroup = groupName });

        Assert.Single(TriviaGroupRegistry.GetGroupPeers(groupName, a));

        await b.DisposeAsync();

        Assert.Empty(TriviaGroupRegistry.GetGroupPeers(groupName, a));
    }
}
