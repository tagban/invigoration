using Invigoration.Scr.Classic;

namespace Invigoration.Scr.LegacyChat;

/// <summary>Something that happened on an SC:R classic connection.</summary>
public abstract record ScrChatEvent
{
    private ScrChatEvent()
    {
    }

    /// <summary>The server confirmed which channel we're in. Only this moves us; sending a join doesn't.</summary>
    public sealed record ChannelEntered(ScrChannel Channel) : ScrChatEvent;

    /// <summary>Channels were added, updated or removed from the list.</summary>
    public sealed record ChannelListChanged(IReadOnlyList<ScrChannelChange> Changes) : ScrChatEvent;

    public sealed record Message(ScrMessage Value) : ScrChatEvent;

    /// <summary>A server call this layer doesn't interpret. It has still been answered.</summary>
    public sealed record Unhandled(uint Service, uint Method, byte[] Body) : ScrChatEvent;

    /// <summary>The server took us out of a channel (LegacyChat.LeftChannel).</summary>
    public sealed record ChannelLeft(ulong ChannelId) : ScrChatEvent;
}

/// <summary>
/// LegacyChat on an already-open, signed-in SC:R classic WebSocket. It turns
/// chat actions into scrambled messages to send, and scrambled messages
/// received into events plus the replies each server call needs. It owns no
/// socket: the caller sends every returned byte array as its own binary
/// WebSocket message, in order.
/// </summary>
public sealed class ScrChatSession
{
    private readonly uint _seed;
    private readonly Dictionary<uint, uint> _pendingMethods = new();

    // Requests are built on whichever thread sends; responses are read on another.
    private readonly Lock _gate = new();
    private uint _token;
    private ScrChannel? _joining;

    /// <param name="seed">The classic connection's scrambling seed from sign-in.</param>
    /// <param name="token">The connection's request counter as sign-in left it. Each request increments it first.</param>
    /// <param name="channelId">The channel sign-in joined, if known.</param>
    public ScrChatSession(uint seed, uint token, ulong channelId)
    {
        _seed = seed;
        _token = token;
        ChannelId = channelId;
    }

    /// <summary>The channel the server last confirmed we're in.</summary>
    public ulong ChannelId { get; private set; }

    public byte[] SendMessage(string text) => Request(LegacyChatRequests.SendMessage(ChannelId, text));

    public byte[] Whisper(string recipient, string text) => Request(LegacyChatRequests.Whisper(ChannelId, recipient, text));

    /// <summary>
    /// Joins (or creates) a channel by name with the chat command. Battle.net confirms it with its
    /// current-channel call, reported as <see cref="ScrChatEvent.ChannelEntered"/>.
    /// </summary>
    public byte[] JoinByName(string channelName)
    {
        _joining = new ScrChannel(0, channelName, channelName, []);
        return Request(LegacyChatRequests.JoinByName(ChannelId, channelName));
    }

    /// <summary>
    /// Joins a listed channel by ID. Battle.net confirms it with a channel-list update that includes
    /// the channel, which is reported as <see cref="ScrChatEvent.ChannelEntered"/>.
    /// </summary>
    public byte[] JoinListedChannel(ulong targetChannelId) =>
        JoinListedChannel(new ScrChannel(targetChannelId, "", "", []));

    /// <summary>
    /// Joins a listed channel. Battle.net confirms it with a channel-list update carrying the
    /// channel, sometimes as a new instance with its own ID, so a match on the name (with members)
    /// counts too. Reported as <see cref="ScrChatEvent.ChannelEntered"/>.
    /// </summary>
    public byte[] JoinListedChannel(ScrChannel target)
    {
        _joining = target;
        return Request(LegacyChatRequests.JoinListedChannel(target.Id));
    }

    public byte[] ListChannels() => Request(LegacyChatRequests.ListChannels());

    public byte[] LeaveChannel() => Request(LegacyChatRequests.LeaveChannel(ChannelId));

    /// <summary>
    /// Any request on the classic connection, from any service: the sign-in and startup calls use
    /// this too. Returns the scrambled message and the token its response will carry.
    /// </summary>
    public (byte[] Message, uint Token) Call(uint service, uint method, byte[] body, byte[]? requestTrace = null)
    {
        uint token;
        lock (_gate)
        {
            token = _token = unchecked(_token + 1);
            _pendingMethods[token] = method;
        }

        var header = new ClassicHeader(service, method, token, ClassicHeader.RequestRouting, 0, ObjectId: 0, IsResponse: false, requestTrace);
        return (ClassicEnvelope.Scramble(ClassicFrame.Encode(header, body), _seed), token);
    }

    /// <summary>Unscrambles one received WebSocket message and handles every RPC in it.</summary>
    public (IReadOnlyList<ScrChatEvent> Events, IReadOnlyList<byte[]> Replies) Receive(ReadOnlySpan<byte> wireMessage) =>
        Receive(wireMessage, out _);

    /// <summary>As <see cref="Receive(ReadOnlySpan{byte})"/>, also handing back the responses to our own requests.</summary>
    public (IReadOnlyList<ScrChatEvent> Events, IReadOnlyList<byte[]> Replies) Receive(ReadOnlySpan<byte> wireMessage, out IReadOnlyList<ClassicRpc> responses)
    {
        var events = new List<ScrChatEvent>();
        var replies = new List<byte[]>();
        var answered = new List<ClassicRpc>();
        responses = answered;
        foreach (var rpc in ClassicFrame.DecodeAll(ClassicEnvelope.Unscramble(wireMessage, _seed)))
        {
            if (rpc.Header.IsResponse)
            {
                // Replies carry no status to check: a refused request shows up as a chat error
                // message or as a join that never gets confirmed.
                lock (_gate)
                {
                    _pendingMethods.Remove(rpc.Header.Token);
                }

                answered.Add(rpc);
                continue;
            }

            replies.Add(ClassicEnvelope.Scramble(LegacyChatCallbacks.Reply(rpc), _seed));
            if (Interpret(rpc) is { } chatEvent)
            {
                events.Add(chatEvent);
            }
        }

        return (events, replies);
    }

    private ScrChatEvent? Interpret(ClassicRpc rpc)
    {
        if (rpc.Header.Service == LegacyChatService.ConnectionHash)
        {
            return null;
        }

        if (rpc.Header.Service != LegacyChatService.Hash)
        {
            return new ScrChatEvent.Unhandled(rpc.Header.Service, rpc.Header.Method, rpc.Body);
        }

        switch (rpc.Header.Method)
        {
            case LegacyChatService.CurrentChannelCallback:
                if (LegacyChatCallbacks.DecodeCurrentChannel(rpc.Body) is not { } channel)
                {
                    return null;
                }

                ChannelId = channel.Id;
                _joining = null;
                return new ScrChatEvent.ChannelEntered(channel);

            case LegacyChatService.LeftChannelCallback:
                var left = LegacyChatCallbacks.DecodeChannelId(rpc.Body);
                if (left == ChannelId)
                {
                    ChannelId = 0;
                }

                return new ScrChatEvent.ChannelLeft(left);

            case LegacyChatService.ChannelListChangedCallback:
                var changes = LegacyChatCallbacks.DecodeChannelListChanges(rpc.Body);
                if (_joining is { } target && changes.FirstOrDefault(c => !c.IsRemoval && IsJoinOf(target, c.Channel)) is { } joined)
                {
                    _joining = null;
                    ChannelId = joined.Channel.Id;
                    return new ScrChatEvent.ChannelEntered(joined.Channel);
                }

                return new ScrChatEvent.ChannelListChanged(changes);

            default:
                if (LegacyChatCallbacks.TryGetMessageKind(rpc.Header.Method, out var kind))
                {
                    return LegacyChatCallbacks.DecodeMessage(kind, rpc.Body) is { } message ? new ScrChatEvent.Message(message) : null;
                }

                return new ScrChatEvent.Unhandled(rpc.Header.Service, rpc.Header.Method, rpc.Body);
        }
    }

    private static bool IsJoinOf(ScrChannel target, ScrChannel channel) =>
        channel.Id == target.Id
        || (channel.Members.Count > 0 && target.DisplayName.Length > 0
            && (channel.DisplayName.Equals(target.DisplayName, StringComparison.OrdinalIgnoreCase)
                || channel.InternalName.Equals(target.InternalName, StringComparison.OrdinalIgnoreCase)));

    private byte[] Request((uint Method, byte[] Body) request) => Call(LegacyChatService.Hash, request.Method, request.Body).Message;
}
