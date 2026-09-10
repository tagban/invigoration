using System.Diagnostics;
using System.Reflection;
using Invigoration.Core.Chat;
using Invigoration.Core.Clan;
using Invigoration.Core.Config;
using Invigoration.Core.Protocol;

namespace Invigoration.Core.Tests;

/// <summary>
/// Regression test for a real, live-diagnosed bug: a mass-join burst locked the whole app (every
/// bot, not just the flooded one) for tens of seconds even after IsJoinBurstActive existed.
/// IsJoinBurstActive alone only stops rank-behavior processing once a *count* of joins crosses a
/// threshold — up to that many tracked+ranked joiners could each still queue a real
/// auto-whisper/-kick/-ban send first, and every send serializes behind BotEngine's single
/// static, shared-across-every-bot flood-protection gate (each holding it across its own
/// ~FloodProtectionDelayMs wait). Confirmed by measurement: before IsChatSendQueueBusy existed,
/// a burst with just 15 pre-threshold tracked+ranked joiners took ~30s (15 * the 2000ms default
/// delay) for the outgoing send queue to settle. This test dispatches a much larger, mixed burst
/// the way the real receive loop actually does — fire-and-forget (see DispatchFireAndForget's
/// remarks), not sequentially awaited like BotEngineJoinBurstTests — and asserts the send queue
/// settles quickly regardless of burst size.
/// </summary>
[Collection("ClanRosterStore")]
public class BotEngineJoinFloodStressTests
{
    private static byte[] BuildJoinFrame(string username) =>
        new PacketWriter()
            .WriteDword((uint)ChatEventType.Join)
            .WriteDword(0)
            .WriteDword(0)
            .WriteDword(0).WriteDword(0).WriteDword(0)
            .WriteNTString(username)
            .WriteNTString("PX2D")
            .ToBncsPacket(BncsPacketId.SID_CHATEVENT);

    private static void DispatchFireAndForget(BotEngine engine, byte[] frame)
    {
        // Mirrors BotEngine.cs line 122: _bncs.PacketReceived += frame => SafeFireAndForget(HandleBncsPacket(frame), ...)
        // i.e. the real receive loop does NOT await each packet's handler before reading the next.
        var method = typeof(BotEngine).GetMethod("HandleBncsPacket", BindingFlags.NonPublic | BindingFlags.Instance)!;
        _ = (Task)method.Invoke(engine, [frame])!;
    }

    [Fact]
    public async Task RealisticBurst_MeasureWallClockTimeToSettle()
    {
        var config = new BotConfig { ClanFeatureEnabled = true, BattlenetServer = "useast.battle.net" };
        await using var engine = new BotEngine(config);
        var rankName = $"rank-{Guid.NewGuid():N}";
        ClanRankStore.Ranks.Add(new ClanRank { Name = rankName, AutoWhisperMessage = "hi", AutoWhisperFrequency = AutoWhisperFrequency.EveryTime });

        // 150 distinct tracked+ranked usernames (simulating a load bot whose accounts were all
        // auto-tracked from a previous run's chat) + 300 plain/untracked joiners — deliberately
        // far larger than the old JoinBurstThreshold (15) so a regression in either gate shows up
        // as a large, unmistakable measured delay rather than something borderline.
        const int trackedCount = 150;
        const int plainCount = 300;

        var trackedNames = new List<string>();
        for (var i = 0; i < trackedCount; i++)
        {
            var name = $"tracked-{i}-{Guid.NewGuid():N}";
            trackedNames.Add(name);
            ClanRosterStore.Members.Add(new ClanMember { Name = name, Rank = rankName });
        }

        var allFrames = new List<byte[]>();
        foreach (var name in trackedNames)
        {
            allFrames.Add(BuildJoinFrame(name));
        }

        for (var i = 0; i < plainCount; i++)
        {
            allFrames.Add(BuildJoinFrame($"plain-{i}-{Guid.NewGuid():N}"));
        }

        try
        {
            var sw = Stopwatch.StartNew();

            // Fire every frame's handler without awaiting any of them individually — this is what
            // actually happens on the wire when many small SID_CHATEVENT frames land in one or a
            // few socket reads.
            foreach (var frame in allFrames)
            {
                DispatchFireAndForget(engine, frame);
            }

            // Now wait for the *outgoing send queue itself* to fully drain — the real symptom the
            // user sees isn't "did the last Task complete" so much as "how long until the app can
            // send/respond normally again." Poll BotEngine's static gate/timestamp via reflection,
            // requiring it to be BOTH in the past AND unchanged for a stability window — a single
            // in-the-past reading isn't enough proof nothing's still queued, since another waiter
            // might not yet have acquired the semaphore and pushed the timestamp further out.
            var nextAllowedField = typeof(BotEngine).GetField("_nextChatSendAllowedUtc", BindingFlags.NonPublic | BindingFlags.Static)!;
            var stableSince = (DateTime?)null;
            DateTime lastObserved = default;
            while (sw.Elapsed < TimeSpan.FromSeconds(60))
            {
                var current = (DateTime)nextAllowedField.GetValue(null)!;
                if (current != lastObserved)
                {
                    lastObserved = current;
                    stableSince = DateTime.UtcNow;
                }

                if (DateTime.UtcNow >= lastObserved && stableSince is { } since && DateTime.UtcNow - since > TimeSpan.FromMilliseconds(300))
                {
                    break;
                }

                await Task.Delay(50);
            }

            sw.Stop();

            Assert.True(sw.Elapsed < TimeSpan.FromSeconds(5),
                $"Outgoing send queue took {sw.Elapsed.TotalSeconds:F1}s to settle after a {trackedCount + plainCount}-join burst ({trackedCount} tracked+ranked, {plainCount} plain).");
        }
        finally
        {
            foreach (var name in trackedNames)
            {
                ClanRosterStore.Members.RemoveAll(m => m.Name == name);
            }

            ClanRankStore.Ranks.RemoveAll(r => r.Name == rankName);
        }
    }
}
