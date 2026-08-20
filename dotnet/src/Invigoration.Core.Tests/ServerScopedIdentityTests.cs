using System.Reflection;
using Invigoration.Core.Chat;
using Invigoration.Core.Clan;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

public class BnetUsernameServerScopingTests
{
    [Theory]
    [InlineData("tagban", "useast.battle.net")]
    [InlineData("tagban", "asia.battle.net")]
    [InlineData("*tagban", "asia.battle.net")]
    public void MatchesOnServer_UnqualifiedEntry_MatchesAnyServer(string speaker, string speakerServer)
    {
        Assert.True(BnetUsername.MatchesOnServer(speaker, "tagban", speakerServer));
    }

    [Fact]
    public void MatchesOnServer_QualifiedEntry_MatchesOnlyItsOwnServer()
    {
        Assert.True(BnetUsername.MatchesOnServer("tagban", "tagban@useast.battle.net", "useast.battle.net"));
    }

    [Fact]
    public void MatchesOnServer_QualifiedEntry_RejectsDifferentServer()
    {
        // The core security scenario: someone else owns "tagban" on Asia — a
        // completely different, unrelated account — and must NOT match an
        // entry deliberately pinned to useast.battle.net.
        Assert.False(BnetUsername.MatchesOnServer("tagban", "tagban@useast.battle.net", "asia.battle.net"));
    }

    [Fact]
    public void MatchesOnServer_ServerComparison_IgnoresBattleNetSuffixAndCase()
    {
        Assert.True(BnetUsername.MatchesOnServer("tagban", "tagban@USEast", "useast.battle.net"));
        Assert.True(BnetUsername.MatchesOnServer("tagban", "tagban@useast.battle.net", "USEast"));
    }

    [Fact]
    public void MatchesOnServer_WrongName_NeverMatchesRegardlessOfServer()
    {
        Assert.False(BnetUsername.MatchesOnServer("someoneelse", "tagban@useast.battle.net", "useast.battle.net"));
    }

    [Fact]
    public void SplitServerQualifier_NoAt_ReturnsNullServer()
    {
        var (name, server) = BnetUsername.SplitServerQualifier("tagban");
        Assert.Equal("tagban", name);
        Assert.Null(server);
    }

    [Fact]
    public void SplitServerQualifier_WithAt_SplitsNameAndServer()
    {
        var (name, server) = BnetUsername.SplitServerQualifier("tagban@useast.battle.net");
        Assert.Equal("tagban", name);
        Assert.Equal("useast.battle.net", server);
    }
}

[Collection("ClanRosterStore")]
public class ClanRosterStoreFindTrustedTests
{
    [Fact]
    public void FindTrusted_QualifiedName_RejectsSameNameOnDifferentServer()
    {
        var name = $"tagban-{Guid.NewGuid():N}";
        ClanRosterStore.Members.Add(new ClanMember { Name = $"{name}@useast.battle.net", Rank = "Leader" });
        try
        {
            var trustedOnEast = ClanRosterStore.FindTrusted(name, "useast.battle.net");
            var trustedOnAsia = ClanRosterStore.FindTrusted(name, "asia.battle.net");

            Assert.NotNull(trustedOnEast);
            Assert.Null(trustedOnAsia);
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name.StartsWith(name));
        }
    }

    [Fact]
    public void FindTrusted_UnqualifiedName_MatchesAnyServer_UnchangedBehavior()
    {
        var name = $"plain-{Guid.NewGuid():N}";
        ClanRosterStore.Members.Add(new ClanMember { Name = name, Rank = "Officer" });
        try
        {
            Assert.NotNull(ClanRosterStore.FindTrusted(name, "useast.battle.net"));
            Assert.NotNull(ClanRosterStore.FindTrusted(name, "asia.battle.net"));
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == name);
        }
    }
}

[Collection("ClanRosterStore")]
public class BotEngineServerScopedAuthorizationTests
{
    private static bool InvokeIsAuthorized(BotEngine engine, string username, string command, string rest = "")
    {
        var method = typeof(BotEngine).GetMethod("IsAuthorized", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (bool)method.Invoke(engine, [username, command, rest])!;
    }

    [Fact]
    public async Task IsAuthorized_QualifiedBotMaster_RejectsSameNameOnDifferentServer()
    {
        var config = new BotConfig { BotMaster = "tagban@useast.battle.net", BattlenetServer = "asia.battle.net" };
        await using var engine = new BotEngine(config);

        // "tagban" on Asia is an unrelated account — must not get bot-master access
        // meant only for the real tagban on useast.battle.net.
        Assert.False(InvokeIsAuthorized(engine, "tagban", "kick"));
    }

    [Fact]
    public async Task IsAuthorized_QualifiedBotMaster_GrantsAccessOnItsOwnServer()
    {
        var config = new BotConfig { BotMaster = "tagban@useast.battle.net", BattlenetServer = "useast.battle.net" };
        await using var engine = new BotEngine(config);

        Assert.True(InvokeIsAuthorized(engine, "tagban", "kick"));
    }

    [Fact]
    public async Task IsAuthorized_UnqualifiedBotMaster_BehavesExactlyAsBefore()
    {
        var config = new BotConfig { BotMaster = "tagban", BattlenetServer = "asia.battle.net" };
        await using var engine = new BotEngine(config);

        Assert.True(InvokeIsAuthorized(engine, "tagban", "kick"));
    }
}
