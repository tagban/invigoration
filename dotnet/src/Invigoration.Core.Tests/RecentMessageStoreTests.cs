using Invigoration.Core.Tracking;

namespace Invigoration.Core.Tests;

/// <summary>Uses guid-suffixed server names per test so parallel/repeated runs against the real, shared, file-backed store never collide — same reasoning as ProtocolUserTrackingStoreTests.</summary>
public class RecentMessageStoreTests
{
    [Fact]
    public void GetRecent_NoMessagesYet_ReturnsEmpty()
    {
        var server = $"server-{Guid.NewGuid():N}";
        Assert.Empty(RecentMessageStore.GetRecent("Hotline", server));
    }

    [Fact]
    public void Append_AddsInOrder()
    {
        var server = $"server-{Guid.NewGuid():N}";

        RecentMessageStore.Append("Hotline", server, "first");
        RecentMessageStore.Append("Hotline", server, "second");

        var recent = RecentMessageStore.GetRecent("Hotline", server);
        Assert.Equal(["first", "second"], recent.Select(m => m.Text));
    }

    [Fact]
    public void Append_BeyondRetentionCount_DropsOldestFirst()
    {
        var server = $"server-{Guid.NewGuid():N}";

        for (var i = 0; i < RecentMessageStore.RetentionCount + 3; i++)
        {
            RecentMessageStore.Append("Hotline", server, $"msg{i}");
        }

        var recent = RecentMessageStore.GetRecent("Hotline", server);
        Assert.Equal(RecentMessageStore.RetentionCount, recent.Count);
        Assert.Equal("msg3", recent[0].Text);
        Assert.Equal($"msg{RecentMessageStore.RetentionCount + 2}", recent[^1].Text);
    }

    [Fact]
    public void Append_SameServerDifferentProtocol_TracksSeparately()
    {
        var server = $"server-{Guid.NewGuid():N}";

        RecentMessageStore.Append("Hotline", server, "hotline message");
        RecentMessageStore.Append("IRC", server, "irc message");

        Assert.Single(RecentMessageStore.GetRecent("Hotline", server));
        Assert.Single(RecentMessageStore.GetRecent("IRC", server));
        Assert.Equal("hotline message", RecentMessageStore.GetRecent("Hotline", server)[0].Text);
    }

    [Fact]
    public void Append_WithTimestamp_PreservesIt()
    {
        var server = $"server-{Guid.NewGuid():N}";
        var ts = DateTimeOffset.UtcNow;

        RecentMessageStore.Append("Hotline", server, "hi", ts);

        Assert.Equal(ts, RecentMessageStore.GetRecent("Hotline", server)[0].TimestampUtc);
    }
}
