using System.Reflection;
using Invigoration.Core.Chat;
using Invigoration.Core.Clan;
using Invigoration.Core.Config;
using Invigoration.Core.Protocol;

namespace Invigoration.Core.Tests;

/// <summary>
/// Covers a real, live-diagnosed bug: a mass-join burst (a load test, or a hostile flood) that
/// includes even a handful of tracked members with an auto-whisper rank previously triggered one
/// real awaited network round-trip per such joiner, serially, in the same pipeline draining the
/// burst itself — the app appeared to hang. IsJoinBurstActive (BotEngine.Bncs.cs) is a channel-
/// wide (not per-username) rolling window that temporarily skips ApplyRankBehaviorsAsync entirely
/// once too many Join/ShowUser events land within a few seconds.
/// </summary>
[Collection("ClanRosterStore")]
public class BotEngineJoinBurstTests
{
    private const int JoinBurstThreshold = 15; // must match BotEngine.Bncs.cs's private constant

    private static byte[] BuildJoinFrame(string username) =>
        new PacketWriter()
            .WriteDword((uint)ChatEventType.Join)
            .WriteDword(0)
            .WriteDword(0)
            .WriteDword(0).WriteDword(0).WriteDword(0)
            .WriteNTString(username)
            .WriteNTString("PX2D")
            .ToBncsPacket(BncsPacketId.SID_CHATEVENT);

    private static Task InvokeHandleChatEvent(BotEngine engine, byte[] frame)
    {
        var method = typeof(BotEngine).GetMethod("HandleBncsChatEventFrame", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (Task)method.Invoke(engine, [frame])!;
    }

    [Fact]
    public async Task JoinBurst_ExceedsThreshold_StopsApplyingRankBehaviorsUntilItSubsides()
    {
        var config = new BotConfig { ClanFeatureEnabled = true, BattlenetServer = "useast.battle.net", FloodProtectionDelayMs = 0 };
        // In a channel, so rank-behavior sends really enter the shared send queue this test is about
        // (an unconnected engine's chat is otherwise held back — see BotEngine.ChatSendBlockedReason).
        await using var engine = BotEngineChatGateTests.MarkInChannel(new BotEngine(config));
        var username = $"test-{Guid.NewGuid():N}";
        var rankName = $"rank-{Guid.NewGuid():N}";
        ClanRankStore.Ranks.Add(new ClanRank { Name = rankName, AutoWhisperMessage = "hi", AutoWhisperFrequency = AutoWhisperFrequency.EveryTime });
        ClanRosterStore.Members.Add(new ClanMember { Name = username, Rank = rankName });
        ClanRosterStore.InvalidateNameIndex();
        try
        {
            // Under the threshold: rank behaviors still run normally, stamping LastAutoWhisperUtc
            // every time ("EveryTime" frequency always re-qualifies).
            for (var i = 0; i < JoinBurstThreshold; i++)
            {
                await InvokeHandleChatEvent(engine, BuildJoinFrame(username));
            }

            var stampBeforeBurst = ClanRosterStore.Find(username)!.LastAutoWhisperUtc;
            Assert.NotNull(stampBeforeBurst);

            // This join pushes the channel-wide count past the threshold — burst-active from here.
            await InvokeHandleChatEvent(engine, BuildJoinFrame(username));

            // Rank behaviors (including the auto-whisper stamp) should NOT have run for this join.
            Assert.Equal(stampBeforeBurst, ClanRosterStore.Find(username)!.LastAutoWhisperUtc);
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == username);
            ClanRosterStore.InvalidateNameIndex();
            ClanRankStore.Ranks.RemoveAll(r => r.Name == rankName);
        }
    }

    [Fact]
    public async Task JoinsUnderThreshold_NeverSuppressRankBehaviors()
    {
        var config = new BotConfig { ClanFeatureEnabled = true, BattlenetServer = "useast.battle.net", FloodProtectionDelayMs = 0 };
        // In a channel, so rank-behavior sends really enter the shared send queue this test is about
        // (an unconnected engine's chat is otherwise held back — see BotEngine.ChatSendBlockedReason).
        await using var engine = BotEngineChatGateTests.MarkInChannel(new BotEngine(config));
        var username = $"test-{Guid.NewGuid():N}";
        var rankName = $"rank-{Guid.NewGuid():N}";
        ClanRankStore.Ranks.Add(new ClanRank { Name = rankName, AutoWhisperMessage = "hi", AutoWhisperFrequency = AutoWhisperFrequency.EveryTime });
        ClanRosterStore.Members.Add(new ClanMember { Name = username, Rank = rankName });
        ClanRosterStore.InvalidateNameIndex();
        try
        {
            await InvokeHandleChatEvent(engine, BuildJoinFrame(username));

            Assert.NotNull(ClanRosterStore.Find(username)!.LastAutoWhisperUtc);
        }
        finally
        {
            ClanRosterStore.Members.RemoveAll(m => m.Name == username);
            ClanRosterStore.InvalidateNameIndex();
            ClanRankStore.Ranks.RemoveAll(r => r.Name == rankName);
        }
    }
}
