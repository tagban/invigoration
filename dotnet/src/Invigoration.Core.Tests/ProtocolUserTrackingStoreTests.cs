using Invigoration.Core.Tracking;

namespace Invigoration.Core.Tests;

/// <summary>Uses guid-suffixed names/servers per test so parallel/repeated runs against the real, shared, file-backed store never collide — same reasoning as the Clan.ClanRosterStore tests.</summary>
public class ProtocolUserTrackingStoreTests
{
    [Fact]
    public void MarkSeen_NewUser_CreatesEntryWithLastSeenSet()
    {
        var name = $"user-{Guid.NewGuid():N}";
        var server = $"server-{Guid.NewGuid():N}";

        var before = DateTime.UtcNow;
        var user = ProtocolUserTrackingStore.MarkSeen(name, "Hotline", server);

        Assert.Equal(name, user.Name);
        Assert.Equal("Hotline", user.Protocol);
        Assert.Equal(server, user.Server);
        Assert.True(user.LastSeenUtc >= before);
        Assert.Equal(0, user.TriviaScore);
    }

    [Fact]
    public void MarkSeen_ExistingUser_UpdatesLastSeenWithoutDuplicating()
    {
        var name = $"user-{Guid.NewGuid():N}";
        var server = $"server-{Guid.NewGuid():N}";

        ProtocolUserTrackingStore.MarkSeen(name, "Hotline", server);
        var firstSeen = ProtocolUserTrackingStore.Find(name, "Hotline", server)!.LastSeenUtc;

        Thread.Sleep(10);
        ProtocolUserTrackingStore.MarkSeen(name, "Hotline", server);
        var second = ProtocolUserTrackingStore.Find(name, "Hotline", server)!;

        Assert.True(second.LastSeenUtc >= firstSeen);
        Assert.Single(ProtocolUserTrackingStore.Users, u => u.Matches(name, "Hotline", server));
    }

    [Fact]
    public void MarkSeen_SameNameDifferentServer_TracksSeparately()
    {
        var name = $"user-{Guid.NewGuid():N}";
        var serverA = $"server-{Guid.NewGuid():N}";
        var serverB = $"server-{Guid.NewGuid():N}";

        ProtocolUserTrackingStore.AddScore(name, "Hotline", serverA, 5);
        ProtocolUserTrackingStore.AddScore(name, "Hotline", serverB, 1);

        Assert.Equal(5, ProtocolUserTrackingStore.Find(name, "Hotline", serverA)!.TriviaScore);
        Assert.Equal(1, ProtocolUserTrackingStore.Find(name, "Hotline", serverB)!.TriviaScore);
    }

    [Fact]
    public void AddScore_AccumulatesAcrossCalls()
    {
        var name = $"user-{Guid.NewGuid():N}";
        var server = $"server-{Guid.NewGuid():N}";

        ProtocolUserTrackingStore.AddScore(name, "Hotline", server, 1.25);
        var total = ProtocolUserTrackingStore.AddScore(name, "Hotline", server, 1.0);

        Assert.Equal(2.25, total);
        Assert.Equal(2.25, ProtocolUserTrackingStore.Find(name, "Hotline", server)!.TriviaScore);
    }

    [Fact]
    public void GetLeaderboard_ScopedToProtocolAndServer_ExcludesOthers()
    {
        var server = $"server-{Guid.NewGuid():N}";
        var otherServer = $"server-{Guid.NewGuid():N}";
        var winner = $"winner-{Guid.NewGuid():N}";
        var elsewhere = $"elsewhere-{Guid.NewGuid():N}";

        ProtocolUserTrackingStore.AddScore(winner, "Hotline", server, 3);
        ProtocolUserTrackingStore.AddScore(elsewhere, "Hotline", otherServer, 99);

        var leaderboard = ProtocolUserTrackingStore.GetLeaderboard("Hotline", server);

        Assert.Contains(leaderboard, u => u.Name == winner);
        Assert.DoesNotContain(leaderboard, u => u.Name == elsewhere);
    }

    [Fact]
    public void QualifiedName_IncludesProtocolAndServer()
    {
        var user = new TrackedUser { Name = "Tagban", Protocol = "Hotline", Server = "bigredh.com:5500" };
        Assert.Equal("Tagban [Hotline:bigredh.com:5500]", user.QualifiedName);
    }
}
