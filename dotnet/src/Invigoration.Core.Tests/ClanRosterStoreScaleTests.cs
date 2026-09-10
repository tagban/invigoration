using System.Diagnostics;
using Invigoration.Core.Clan;

namespace Invigoration.Core.Tests;

/// <summary>
/// Regression test for lag reappearing specifically at scale (reported: "still getting lagged
/// around 1000 mark") even after the per-event send-pileup fixes (IsJoinBurstActive,
/// IsChatSendQueueBusy) resolved the earlier burst-rate hang. Root cause: Find/FindTrusted did a
/// plain Members.FirstOrDefault(m => m.Matches(...)) linear scan — fine for a roster of a few
/// dozen formal clan members, but every Join/ShowUser/UserFlags event calls this (RecordProductSeen,
/// ApplyRankBehaviorsAsync), and repeated load testing naturally grows the roster (RecordSeen
/// auto-tracks "everyone who's ever spoken") into the thousands over time — at that size, the
/// per-lookup scan cost itself becomes the bottleneck, independent of burst *rate*. NameIndex
/// (a Dictionary keyed by normalized name/alias) makes this O(1) regardless of roster size.
/// </summary>
[Collection("ClanRosterStore")]
public class ClanRosterStoreScaleTests
{
    [Fact]
    public void Find_WithLargeRoster_StaysFast()
    {
        const int rosterSize = 2000;
        var names = new List<string>(rosterSize);
        for (var i = 0; i < rosterSize; i++)
        {
            var name = $"seen-{i}-{Guid.NewGuid():N}";
            names.Add(name);
            ClanRosterStore.Members.Add(new ClanMember { Name = name, Aliases = [$"alt-{name}"] });
        }

        ClanRosterStore.InvalidateNameIndex();

        try
        {
            // Force one index build up front (simulates the burst already having warmed it),
            // then measure a large number of subsequent lookups — a mix of hits (by primary name
            // and by alias) and misses (never-seen usernames, the common case during a real
            // mass-join flood of mostly-unfamiliar accounts).
            ClanRosterStore.Find(names[0]);

            var sw = Stopwatch.StartNew();
            var hits = 0;
            for (var i = 0; i < rosterSize; i++)
            {
                if (ClanRosterStore.Find(names[i]) is not null)
                {
                    hits++;
                }

                if (ClanRosterStore.Find("alt-" + names[i]) is not null)
                {
                    hits++;
                }

                ClanRosterStore.Find($"never-seen-{i}-{Guid.NewGuid():N}");
            }

            sw.Stop();

            Assert.Equal(rosterSize * 2, hits);
            // A pre-fix linear scan here is O(rosterSize) per lookup * 3*rosterSize lookups =
            // O(rosterSize^2) = 12,000,000 comparisons for this test alone; with the index this
            // is dominated by the single rebuild plus O(1) dictionary lookups. Generous ceiling
            // (real hardware, Debug build) — the point is "milliseconds, not seconds."
            Assert.True(sw.Elapsed < TimeSpan.FromSeconds(2),
                $"{rosterSize * 3} lookups against a {rosterSize}-member roster took {sw.Elapsed.TotalMilliseconds:F0}ms.");
        }
        finally
        {
            foreach (var name in names)
            {
                ClanRosterStore.Members.RemoveAll(m => m.Name == name);
            }

            ClanRosterStore.InvalidateNameIndex();
        }
    }

    [Fact]
    public void FindTrusted_ServerQualifiedAlias_StillResolvesWithLargeRoster()
    {
        const int rosterSize = 1500;
        var names = new List<string>(rosterSize);
        for (var i = 0; i < rosterSize; i++)
        {
            var name = $"seen-{i}-{Guid.NewGuid():N}";
            names.Add(name);
            ClanRosterStore.Members.Add(new ClanMember { Name = name });
        }

        var target = $"target-{Guid.NewGuid():N}";
        ClanRosterStore.Members.Add(new ClanMember
        {
            Name = $"{target}@useast.battle.net",
            Rank = "Officer",
        });
        ClanRosterStore.InvalidateNameIndex();

        try
        {
            var found = ClanRosterStore.FindTrusted(target, "useast.battle.net");
            Assert.NotNull(found);
            Assert.Equal("Officer", found!.Rank);

            // A different server must NOT match a server-qualified entry, even via the index's
            // name-only bucketing (the actual server check still has to run on the candidates).
            Assert.Null(ClanRosterStore.FindTrusted(target, "asia.battle.net"));
        }
        finally
        {
            foreach (var name in names)
            {
                ClanRosterStore.Members.RemoveAll(m => m.Name == name);
            }

            ClanRosterStore.Members.RemoveAll(m => m.Name.StartsWith(target));
            ClanRosterStore.InvalidateNameIndex();
        }
    }
}
