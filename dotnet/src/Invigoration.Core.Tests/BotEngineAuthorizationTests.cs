using System.Reflection;
using Invigoration.Core.Clan;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>
/// IsAuthorized is private (no public entry point exercises it without
/// simulating a raw BNCS chat frame end-to-end), so these call it via
/// reflection — it's pure given Config/ClanRosterStore, so that's safe and
/// far cheaper than standing up a fake connection just to test permission logic.
/// </summary>
[Collection("ClanRosterStore")]
public class BotEngineAuthorizationTests
{
    private static bool InvokeIsAuthorized(BotEngine engine, string username, string command, string rest = "")
    {
        var method = typeof(BotEngine).GetMethod("IsAuthorized", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (bool)method.Invoke(engine, [username, command, rest])!;
    }

    private static (BotEngine Engine, string MemberName) CreateEngineWithTrackedMember(string rank, string bannedRank = "Banned")
    {
        var config = new BotConfig { BotMaster = "TheMaster", BannedRank = bannedRank };
        var engine = new BotEngine(config);
        var name = $"test-{Guid.NewGuid():N}";
        ClanRosterStore.Members.Add(new ClanMember { Name = name, Rank = rank });
        return (engine, name);
    }

    [Fact]
    public async Task IsAuthorized_BotMaster_AlwaysTrue()
    {
        var config = new BotConfig { BotMaster = "TheMaster" };
        await using var engine = new BotEngine(config);

        Assert.True(InvokeIsAuthorized(engine, "TheMaster", "kick"));
    }

    [Fact]
    public async Task IsAuthorized_BannedRank_BlocksEverythingIncludingTriviaScore()
    {
        var (engine, name) = CreateEngineWithTrackedMember("Banned");
        await using var _ = engine;
        try
        {
            Assert.False(InvokeIsAuthorized(engine, name, "trivia", "score"));
            Assert.False(InvokeIsAuthorized(engine, name, "kick"));
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == name);
        }
    }

    [Fact]
    public async Task IsAuthorized_UntrackedUser_TriviaScoreAndCategoriesAreOpen()
    {
        var config = new BotConfig { BotMaster = "TheMaster" };
        await using var engine = new BotEngine(config);
        var name = $"untracked-{Guid.NewGuid():N}";

        Assert.True(InvokeIsAuthorized(engine, name, "trivia", "score"));
        Assert.True(InvokeIsAuthorized(engine, name, "trivia", "categories"));
    }

    [Fact]
    public async Task IsAuthorized_UntrackedUser_AdminCommandsAreNotOpen()
    {
        var config = new BotConfig { BotMaster = "TheMaster" };
        await using var engine = new BotEngine(config);
        var name = $"untracked-{Guid.NewGuid():N}";

        Assert.False(InvokeIsAuthorized(engine, name, "kick"));
        Assert.False(InvokeIsAuthorized(engine, name, "trivia", "on"));
    }

    [Fact]
    public async Task IsAuthorized_RankWithNoAllowedCommands_TriviaRoundControlStillRequiresGrant()
    {
        var rankName = $"rank-{Guid.NewGuid():N}";
        ClanRankStore.Ranks.Add(new ClanRank { Name = rankName });
        var (engine, name) = CreateEngineWithTrackedMember(rankName);
        await using var _ = engine;
        try
        {
            // A rank with no AllowedCommands doesn't grant "trivia on"/"off" (round control).
            Assert.False(InvokeIsAuthorized(engine, name, "trivia", "on"));
            // But score stays open regardless of rank, as long as they're not banned.
            Assert.True(InvokeIsAuthorized(engine, name, "trivia", "score"));
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == name);
            ClanRankStore.Ranks.RemoveAll(r => r.Name == rankName);
        }
    }

    [Fact]
    public async Task IsAuthorized_RankGrantsCommand_AuthorizesResolvedCanonicalName()
    {
        var rankName = $"rank-{Guid.NewGuid():N}";
        ClanRankStore.Ranks.Add(new ClanRank { Name = rankName, AllowedCommands = ["kick"] });
        var (engine, name) = CreateEngineWithTrackedMember(rankName);
        await using var _ = engine;
        try
        {
            Assert.True(InvokeIsAuthorized(engine, name, "kick"));
            Assert.False(InvokeIsAuthorized(engine, name, "ban"));
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == name);
            ClanRankStore.Ranks.RemoveAll(r => r.Name == rankName);
        }
    }
}

[Collection("ClanRosterStore")]
public class ClanRosterStoreAutoRegistrationTests
{
    [Fact]
    public void RecordSeen_UntrackedUser_WithDefaultRank_CreatesMember()
    {
        var name = $"newcomer-{Guid.NewGuid():N}";
        try
        {
            ClanRosterStore.RecordSeen(name, "Trivia Participant");

            var member = ClanRosterStore.Find(name);
            Assert.NotNull(member);
            Assert.Equal("Trivia Participant", member!.Rank);
            Assert.NotNull(member.LastSeenUtc);
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == name);
            ClanRosterStore.Save();
        }
    }

    [Fact]
    public void RecordSeen_UntrackedUser_NullDefaultRank_StaysUntracked()
    {
        var name = $"nobody-{Guid.NewGuid():N}";

        ClanRosterStore.RecordSeen(name, null);

        Assert.Null(ClanRosterStore.Find(name));
    }

    [Fact]
    public void RecordSeen_ExistingMember_KeepsRankEvenWithDefaultRankPassed()
    {
        var name = $"existing-{Guid.NewGuid():N}";
        ClanRosterStore.Members.Add(new ClanMember { Name = name, Rank = "Officer" });
        try
        {
            ClanRosterStore.RecordSeen(name, "Trivia Participant");

            Assert.Equal("Officer", ClanRosterStore.Find(name)!.Rank);
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == name);
            ClanRosterStore.Save();
        }
    }
}
