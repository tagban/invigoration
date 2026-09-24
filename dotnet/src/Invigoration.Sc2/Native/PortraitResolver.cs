namespace Invigoration.Sc2.Native;

// Ported from ncarrillo/superiority (MIT): core/src/games/sc2/chat/session.rs's
// ProfileResolver (enqueue, pump and receive).

/// <summary>
/// Looks up StarCraft II portraits by reading profiles, one profile read per address per
/// session, with every answer cached, misses included. Pure state: it builds no records
/// and sends nothing. Per session:
/// <list type="number">
/// <item><see cref="Enqueue"/> the profile addresses of chat members and friends that
/// have no presence avatar.</item>
/// <item><see cref="NextRequests"/> after every record (and on a timer), sending
/// <see cref="ChatCommands.ProfileReadRequest(uint, PlayerTarget.ProfileRecordAddress)"/>
/// for each pair it returns.</item>
/// <item><see cref="Complete"/> with every <see cref="NativeChatRecord.ProfileRead"/>.</item>
/// <item><see cref="For(PlayerTarget.ProfileRecordAddress)"/> to read the result.</item>
/// </list>
/// Requests are limited as upstream does: at most <see cref="MaxInFlight"/> unanswered,
/// and a token bucket refilling <see cref="RatePerSecond"/> a second up to
/// <see cref="Burst"/>. Not thread-safe; use it from the receive loop.
/// </summary>
public sealed class PortraitResolver
{
    public const int MaxInFlight = 16;
    public const double RatePerSecond = 40;
    public const double Burst = 20;

    /// <summary>
    /// A request still unanswered after this long counts as a miss, so a lost answer
    /// can't hold one of the <see cref="MaxInFlight"/> slots forever. Upstream has no
    /// such limit.
    /// </summary>
    public static readonly TimeSpan PendingTimeout = TimeSpan.FromSeconds(30);

    private readonly Dictionary<PlayerTarget.ProfileRecordAddress, Sc2Portrait?> _resolved = [];
    private readonly HashSet<PlayerTarget.ProfileRecordAddress> _queued = [];
    private readonly Queue<PlayerTarget.ProfileRecordAddress> _queue = new();
    private readonly Dictionary<uint, Pending> _pending = [];
    private uint _nextRequestId;
    private double _tokens;
    private DateTimeOffset? _lastRefill;

    public PortraitResolver(uint firstRequestId = 0)
    {
        _nextRequestId = firstRequestId;
    }

    /// <summary>Requests sent and not yet answered.</summary>
    public int InFlight => _pending.Count;

    /// <summary>Addresses waiting for a request slot.</summary>
    public int Queued => _queue.Count;

    /// <summary>
    /// Queues a profile read of <paramref name="address"/>. Returns false, doing nothing,
    /// when it's already resolved, queued or in flight.
    /// </summary>
    public bool Enqueue(PlayerTarget.ProfileRecordAddress address)
    {
        if (_resolved.ContainsKey(address) || _queued.Contains(address) || _pending.Values.Any(p => p.Address == address))
        {
            return false;
        }

        _queued.Add(address);
        _queue.Enqueue(address);
        return true;
    }

    /// <summary>
    /// Queues the profile of a chat member (by <see cref="MembershipChange.Join.PresenceId"/>)
    /// when their presence has a profile address but no avatar of its own, as upstream does.
    /// </summary>
    public bool EnqueueMember(PresenceTracker presence, uint presenceId) =>
        presence.AvatarFor(presenceId) is null && presence.ProfileFor(presenceId) is { } address && Enqueue(address);

    /// <summary>Queues a friend's profile (from the friends list, else their presence), as upstream does whether or not their presence has an avatar.</summary>
    public bool EnqueueFriend(PresenceTracker presence, FriendEntry friend) =>
        presence.ProfileFor(friend) is { } address && Enqueue(address);

    /// <summary>
    /// The profile reads to send now, each with the request id its answers will carry.
    /// Also gives up (as misses) on requests unanswered for <see cref="PendingTimeout"/>.
    /// The result is complete when returned; state has already moved on.
    /// </summary>
    public IEnumerable<(uint RequestId, PlayerTarget.ProfileRecordAddress Address)> NextRequests(DateTimeOffset now)
    {
        foreach (var (requestId, pending) in _pending.Where(p => now - p.Value.SentAt >= PendingTimeout).ToList())
        {
            _pending.Remove(requestId);
            _resolved[pending.Address] = null;
        }

        if (_lastRefill is { } last)
        {
            var elapsed = Math.Max(0, (now - last).TotalSeconds);
            _tokens = Math.Min(Burst, _tokens + (elapsed * RatePerSecond));
        }
        else
        {
            _tokens = Burst;
        }

        _lastRefill = now;
        var requests = new List<(uint, PlayerTarget.ProfileRecordAddress)>();
        while (_pending.Count < MaxInFlight && _tokens >= 1 && _queue.TryDequeue(out var address))
        {
            _tokens -= 1;
            _queued.Remove(address);
            var requestId = _nextRequestId;
            _nextRequestId = unchecked(_nextRequestId + 1);
            _pending[requestId] = new Pending(address, now);
            requests.Add((requestId, address));
        }

        return requests;
    }

    /// <summary>
    /// Takes one answer to a profile read. Start with no packets, Cache and Failure are
    /// misses; Start with packets waits for that many Blocks, whose bytes are gathered
    /// until one holds the portrait or all have arrived. A Block with no Start before it
    /// settles the request by itself, as upstream does. Returns true when this answer
    /// settled an address (hit or miss); answers to requests it didn't make are ignored.
    /// </summary>
    public bool Complete(ProfileReadRecord record)
    {
        if (!_pending.TryGetValue(record.RequestId, out var pending))
        {
            return false;
        }

        switch (record.Kind)
        {
            case ProfileReadKind.Start when record.PacketCount > 0:
                pending.Expected = record.PacketCount;
                return false;
            case ProfileReadKind.Block:
                pending.Data.AddRange(record.Block ?? []);
                pending.Received++;
                var data = pending.Data.ToArray();
                if (Sc2PortraitBlock.TryReadUnlockableId(data, out _)
                    || pending.Expected is not { } expected
                    || pending.Received >= expected)
                {
                    return Settle(record.RequestId, pending.Address, Sc2PortraitBlock.PortraitIn(data));
                }

                return false;
            default:
                return Settle(record.RequestId, pending.Address, null);
        }
    }

    /// <summary>The portrait read from <paramref name="address"/>'s profile, or null when it isn't known (yet, or at all).</summary>
    public Sc2Portrait? For(PlayerTarget.ProfileRecordAddress address) =>
        _resolved.TryGetValue(address, out var portrait) ? portrait : null;

    /// <summary>Whether a read of <paramref name="address"/> has finished, found or not.</summary>
    public bool IsResolved(PlayerTarget.ProfileRecordAddress address) => _resolved.ContainsKey(address);

    /// <summary>
    /// A chat member's portrait as upstream shows it: their presence avatar, else what
    /// their profile read found.
    /// </summary>
    public Sc2Portrait? For(PresenceTracker presence, uint presenceId) =>
        presence.AvatarFor(presenceId) ?? (presence.ProfileFor(presenceId) is { } address ? For(address) : null);

    /// <summary>A friend's portrait: what their profile read found, else their presence avatar.</summary>
    public Sc2Portrait? For(PresenceTracker presence, FriendEntry friend) =>
        (presence.ProfileFor(friend) is { } address ? For(address) : null) ?? presence.AvatarFor(friend);

    private bool Settle(uint requestId, PlayerTarget.ProfileRecordAddress address, Sc2Portrait? portrait)
    {
        _pending.Remove(requestId);
        _resolved[address] = portrait;
        return true;
    }

    private sealed class Pending(PlayerTarget.ProfileRecordAddress address, DateTimeOffset sentAt)
    {
        public PlayerTarget.ProfileRecordAddress Address { get; } = address;

        public DateTimeOffset SentAt { get; } = sentAt;

        public uint? Expected { get; set; }

        public int Received { get; set; }

        public List<byte> Data { get; } = [];
    }
}
