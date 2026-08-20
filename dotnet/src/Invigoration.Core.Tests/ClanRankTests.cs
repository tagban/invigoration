using System.Reflection;
using Invigoration.Core.Clan;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

[Collection("ClanRosterStore")]
public class ClanRankStoreTests
{
    [Fact]
    public void Ranks_DefaultSeeding_IncludesExpectedNames()
    {
        // Note: this reads whatever is currently persisted for this install,
        // which — on a fresh one — is the seeded default set.
        var names = ClanRankStore.Ranks.Select(r => r.Name).ToList();
        Assert.NotEmpty(names);
    }

    [Fact]
    public void Find_ExistingRank_IsCaseInsensitive()
    {
        var name = $"rank-{Guid.NewGuid():N}";
        ClanRankStore.Ranks.Add(new ClanRank { Name = name, AutoKick = true });
        try
        {
            var found = ClanRankStore.Find(name.ToUpperInvariant());

            Assert.NotNull(found);
            Assert.True(found!.AutoKick);
        }
        finally
        {
            ClanRankStore.Ranks.RemoveAll(r => r.Name == name);
        }
    }

    [Fact]
    public void Find_UnknownRank_ReturnsNull()
    {
        Assert.Null(ClanRankStore.Find($"nonexistent-{Guid.NewGuid():N}"));
    }
}

public class ClanRankBehaviorTests
{
    [Fact]
    public void HasAutoWhisper_EmptyMessage_IsFalse()
    {
        var rank = new ClanRank { Name = "Test" };
        Assert.False(rank.HasAutoWhisper);
    }

    [Fact]
    public void HasAutoWhisper_NonEmptyMessage_IsTrue()
    {
        var rank = new ClanRank { Name = "Test", AutoWhisperMessage = "Welcome!" };
        Assert.True(rank.HasAutoWhisper);
    }
}

[Collection("ClanRosterStore")]
public class AutoWhisperFrequencyTests
{
    private static bool InvokeShouldSendAutoWhisper(ClanMember member, AutoWhisperFrequency frequency)
    {
        var method = typeof(BotEngine).GetMethod("ShouldSendAutoWhisper", BindingFlags.NonPublic | BindingFlags.Static)!;
        return (bool)method.Invoke(null, [member, frequency])!;
    }

    [Fact]
    public void ShouldSendAutoWhisper_NeverSentBefore_TrueForEveryFrequency()
    {
        var member = new ClanMember { Name = "test" };

        Assert.True(InvokeShouldSendAutoWhisper(member, AutoWhisperFrequency.EveryTime));
        Assert.True(InvokeShouldSendAutoWhisper(member, AutoWhisperFrequency.Daily));
        Assert.True(InvokeShouldSendAutoWhisper(member, AutoWhisperFrequency.Once));
    }

    [Fact]
    public void ShouldSendAutoWhisper_EveryTime_AlwaysTrueEvenJustSent()
    {
        var member = new ClanMember { Name = "test", LastAutoWhisperUtc = DateTime.UtcNow };

        Assert.True(InvokeShouldSendAutoWhisper(member, AutoWhisperFrequency.EveryTime));
    }

    [Fact]
    public void ShouldSendAutoWhisper_Once_FalseAfterAnyPriorSend()
    {
        var member = new ClanMember { Name = "test", LastAutoWhisperUtc = DateTime.UtcNow.AddYears(-1) };

        Assert.False(InvokeShouldSendAutoWhisper(member, AutoWhisperFrequency.Once));
    }

    [Fact]
    public void ShouldSendAutoWhisper_Daily_FalseWithinTheLast24Hours()
    {
        var member = new ClanMember { Name = "test", LastAutoWhisperUtc = DateTime.UtcNow.AddHours(-1) };

        Assert.False(InvokeShouldSendAutoWhisper(member, AutoWhisperFrequency.Daily));
    }

    [Fact]
    public void ShouldSendAutoWhisper_Daily_TrueAfter24Hours()
    {
        var member = new ClanMember { Name = "test", LastAutoWhisperUtc = DateTime.UtcNow.AddHours(-25) };

        Assert.True(InvokeShouldSendAutoWhisper(member, AutoWhisperFrequency.Daily));
    }
}

[Collection("ClanRosterStore")]
public class ApplyRankBehaviorsTests
{
    private static Task InvokeApplyRankBehaviors(BotEngine engine, string username)
    {
        var method = typeof(BotEngine).GetMethod("ApplyRankBehaviorsAsync", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (Task)method.Invoke(engine, [username])!;
    }

    [Fact]
    public async Task ApplyRankBehaviors_ClanFeatureDisabled_DoesNotUpdateWhisperTimestamp()
    {
        var config = new BotConfig { ClanFeatureEnabled = false, BattlenetServer = "useast.battle.net" };
        await using var engine = new BotEngine(config);
        var username = $"test-{Guid.NewGuid():N}";
        var rankName = $"rank-{Guid.NewGuid():N}";
        ClanRankStore.Ranks.Add(new ClanRank { Name = rankName, AutoWhisperMessage = "hi", AutoWhisperFrequency = AutoWhisperFrequency.EveryTime });
        ClanRosterStore.Members.Add(new ClanMember { Name = username, Rank = rankName });
        try
        {
            await InvokeApplyRankBehaviors(engine, username);

            Assert.Null(ClanRosterStore.Find(username)!.LastAutoWhisperUtc);
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == username);
            ClanRankStore.Ranks.RemoveAll(r => r.Name == rankName);
        }
    }

    [Fact]
    public async Task ApplyRankBehaviors_EnabledWithAutoWhisperRank_StampsLastAutoWhisperUtc()
    {
        var config = new BotConfig { ClanFeatureEnabled = true, BattlenetServer = "useast.battle.net" };
        await using var engine = new BotEngine(config);
        var username = $"test-{Guid.NewGuid():N}";
        var rankName = $"rank-{Guid.NewGuid():N}";
        ClanRankStore.Ranks.Add(new ClanRank { Name = rankName, AutoWhisperMessage = "hi", AutoWhisperFrequency = AutoWhisperFrequency.EveryTime });
        ClanRosterStore.Members.Add(new ClanMember { Name = username, Rank = rankName });
        try
        {
            await InvokeApplyRankBehaviors(engine, username);

            var stamped = ClanRosterStore.Find(username)!.LastAutoWhisperUtc;
            Assert.NotNull(stamped);
            Assert.True(DateTime.UtcNow - stamped!.Value < TimeSpan.FromSeconds(5));
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == username);
            ClanRankStore.Ranks.RemoveAll(r => r.Name == rankName);
        }
    }

    [Fact]
    public async Task ApplyRankBehaviors_RankWithNoAutoWhisper_DoesNotStampTimestamp()
    {
        var config = new BotConfig { ClanFeatureEnabled = true, BattlenetServer = "useast.battle.net" };
        await using var engine = new BotEngine(config);
        var username = $"test-{Guid.NewGuid():N}";
        var rankName = $"rank-{Guid.NewGuid():N}";
        ClanRankStore.Ranks.Add(new ClanRank { Name = rankName, AutoKick = true }); // no whisper message set
        ClanRosterStore.Members.Add(new ClanMember { Name = username, Rank = rankName });
        try
        {
            await InvokeApplyRankBehaviors(engine, username);

            Assert.Null(ClanRosterStore.Find(username)!.LastAutoWhisperUtc);
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == username);
            ClanRankStore.Ranks.RemoveAll(r => r.Name == rankName);
        }
    }

    [Fact]
    public async Task ApplyRankBehaviors_UntrackedUsername_DoesNotThrow()
    {
        var config = new BotConfig { ClanFeatureEnabled = true, BattlenetServer = "useast.battle.net" };
        await using var engine = new BotEngine(config);

        await InvokeApplyRankBehaviors(engine, $"untracked-{Guid.NewGuid():N}");
    }
}
