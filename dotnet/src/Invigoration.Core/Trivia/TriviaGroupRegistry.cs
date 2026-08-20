namespace Invigoration.Core.Trivia;

/// <summary>
/// Lets multiple bots share one trivia round via BotConfig.TriviaGroup — e.g.
/// one bot on Warcraft II and another on StarCraft II, run by the same
/// person, both relaying the same question to their own channel and sharing
/// one player pool/score flow, instead of each running an independent round.
/// Two responsibilities:
/// 1. One shared <see cref="TriviaSession"/> per group name, so "who's
///    joined" and "what's the current question" are the same state no
///    matter which group member's channel a chat message arrives on.
/// 2. Tracking every live BotEngine so a round's owning engine can find its
///    group peers to relay outgoing messages to. Membership is resolved by
///    each engine's *current* Config.TriviaGroup at call time (not
///    snapshotted at registration), since Config can be replaced wholesale
///    after an edit in the Config window.
/// </summary>
public static class TriviaGroupRegistry
{
    private static readonly Lock SyncRoot = new();
    private static readonly Dictionary<string, TriviaSession> Sessions = new(StringComparer.OrdinalIgnoreCase);
    private static readonly List<BotEngine> LiveEngines = [];

    public static void RegisterEngine(BotEngine engine)
    {
        lock (SyncRoot)
        {
            if (!LiveEngines.Contains(engine))
            {
                LiveEngines.Add(engine);
            }
        }
    }

    public static void UnregisterEngine(BotEngine engine)
    {
        lock (SyncRoot)
        {
            LiveEngines.Remove(engine);
        }
    }

    /// <summary>The shared session for this group name, created on first use and reused by every engine naming the same group.</summary>
    public static TriviaSession GetSession(string groupName)
    {
        lock (SyncRoot)
        {
            if (!Sessions.TryGetValue(groupName, out var session))
            {
                session = new TriviaSession();
                Sessions[groupName] = session;
            }

            return session;
        }
    }

    /// <summary>Every other live engine currently naming this same group (case-insensitive), excluding the caller.</summary>
    public static IReadOnlyList<BotEngine> GetGroupPeers(string groupName, BotEngine self)
    {
        lock (SyncRoot)
        {
            return LiveEngines.Where(e => !ReferenceEquals(e, self) && groupName.Equals(e.Config.TriviaGroup, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }
    }
}
