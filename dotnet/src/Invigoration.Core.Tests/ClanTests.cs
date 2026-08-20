using Invigoration.Core.Chat;
using Invigoration.Core.Clan;

namespace Invigoration.Core.Tests;

public class BnetUsernameTests
{
    [Theory]
    [InlineData("Tagban", "Tagban")]
    [InlineData("*Tagban", "Tagban")]
    [InlineData("Tagban", "*Tagban")]
    [InlineData("*Tagban", "*Tagban")]
    [InlineData("TAGBAN", "tagban")]
    public void Equals_IgnoresLeadingAsteriskAndCase(string a, string b)
    {
        Assert.True(BnetUsername.Equals(a, b));
    }

    [Fact]
    public void Equals_DifferentNames_NotEqual()
    {
        Assert.False(BnetUsername.Equals("Tagban", "*SomeoneElse"));
    }

    [Fact]
    public void Normalize_StripsOnlyLeadingAsterisk()
    {
        Assert.Equal("Tagban", BnetUsername.Normalize("*Tagban"));
        Assert.Equal("Tagban", BnetUsername.Normalize("Tagban"));
    }
}

public class ClanMemberTests
{
    [Fact]
    public void Matches_PrimaryName_TolerantOfInGameAsterisk()
    {
        var member = new ClanMember { Name = "Tagban" };

        Assert.True(member.Matches("*Tagban"));
    }

    [Fact]
    public void Matches_Alias_TolerantOfInGameAsterisk()
    {
        var member = new ClanMember { Name = "Tagban", Aliases = ["AltAccount"] };

        Assert.True(member.Matches("*AltAccount"));
    }

    [Fact]
    public void Matches_UnrelatedName_ReturnsFalse()
    {
        var member = new ClanMember { Name = "Tagban", Aliases = ["AltAccount"] };

        Assert.False(member.Matches("SomeoneElse"));
    }
}

/// <summary>In the shared "ClanRosterStore" xUnit collection (see other files touching ClanRosterStore) so these never run concurrently with each other — direct unsynchronized Members.Add/RemoveAll in test code races against the same static list otherwise.</summary>
[Collection("ClanRosterStore")]
public class ClanRosterStoreTests
{
    [Fact]
    public void Find_ResolvesByAliasRegardlessOfAsterisk()
    {
        var name = $"test-{Guid.NewGuid():N}";
        var alias = $"alt-{Guid.NewGuid():N}";
        ClanRosterStore.Members.Add(new ClanMember { Name = name, Rank = "Officer", Aliases = [alias] });
        try
        {
            var found = ClanRosterStore.Find("*" + alias);

            Assert.NotNull(found);
            Assert.Equal(name, found!.Name);
            Assert.Equal("Officer", found.Rank);
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == name);
        }
    }

    /// <summary>Save() persists to the real on-disk roster, so this cleans up by re-saving without the test member rather than just removing it in-memory, to avoid leaving a phantom entry in the user's actual clan-members.json.</summary>
    [Fact]
    public void RecordSeen_TrackedMember_StampsLastSeenUtc()
    {
        var name = $"test-{Guid.NewGuid():N}";
        ClanRosterStore.Members.Add(new ClanMember { Name = name });
        try
        {
            ClanRosterStore.RecordSeen(name);

            var member = ClanRosterStore.Find(name);
            Assert.NotNull(member!.LastSeenUtc);
            Assert.True(DateTime.UtcNow - member.LastSeenUtc!.Value < TimeSpan.FromSeconds(5));
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == name);
            ClanRosterStore.Save();
        }
    }

    [Fact]
    public void RecordSeen_UntrackedUsername_DoesNotThrow()
    {
        ClanRosterStore.RecordSeen($"untracked-{Guid.NewGuid():N}");
    }

    [Fact]
    public void RecordSeen_WithProductAndServer_StampsBoth()
    {
        var name = $"test-{Guid.NewGuid():N}";
        ClanRosterStore.Members.Add(new ClanMember { Name = name });
        try
        {
            ClanRosterStore.RecordSeen(name, product: "PX2D", server: "useast.battle.net");

            var member = ClanRosterStore.Find(name);
            Assert.Equal("PX2D", member!.LastSeenProduct);
            Assert.Equal("useast.battle.net", member.LastSeenServer);
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == name);
            ClanRosterStore.Save();
        }
    }

    [Fact]
    public void RecordSeen_NewMember_IsNotMarkedAsFormalClanMember()
    {
        var name = $"test-{Guid.NewGuid():N}";
        try
        {
            ClanRosterStore.RecordSeen(name, defaultRankIfNew: "Trivia Participant");

            var member = ClanRosterStore.Find(name);
            Assert.NotNull(member);
            Assert.False(member!.IsClanMember);
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == name);
            ClanRosterStore.Save();
        }
    }

    [Fact]
    public void RecordProductSeen_TrackedMember_UpdatesProductAndServer()
    {
        var name = $"test-{Guid.NewGuid():N}";
        ClanRosterStore.Members.Add(new ClanMember { Name = name });
        try
        {
            ClanRosterStore.RecordProductSeen(name, "3RAW", "useast.battle.net");

            var member = ClanRosterStore.Find(name);
            Assert.Equal("3RAW", member!.LastSeenProduct);
            Assert.Equal("useast.battle.net", member.LastSeenServer);
            Assert.Null(member.LastSeenUtc); // presence alone shouldn't stamp LastSeenUtc — only actual chat does (RecordSeen)
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == name);
            ClanRosterStore.Save();
        }
    }

    [Fact]
    public void RecordProductSeen_UntrackedUsername_DoesNotCreateEntry()
    {
        var name = $"untracked-{Guid.NewGuid():N}";

        ClanRosterStore.RecordProductSeen(name, "3RAW", "useast.battle.net");

        Assert.Null(ClanRosterStore.Find(name));
    }
}
