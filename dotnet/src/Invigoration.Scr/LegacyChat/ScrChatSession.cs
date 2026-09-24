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

    /// <summary>A reply to one of our requests reported an error.</summary>
    public sealed record RequestFailed(uint Method, uint Token, uint Status) : ScrChatEvent;

    /// <summary>A server call this layer doesn't interpret. It has still been answered.</summary>
    public sealed record Unhandled(uint Service, uint Method) : ScrChatEvent;
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
    private uint _token;

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

    public byte[] JoinByName(string channelName) => Request(LegacyChatRequests.JoinByName(ChannelId, channelName));

    public byte[] JoinListedChannel(ulong targetChannelId) => Request(LegacyChatRequests.JoinListedChannel(targetChannelId));

    public byte[] ListChannels() => Request(LegacyChatRequests.ListChannels());

    public byte[] LeaveChannel() => Request(LegacyChatRequests.LeaveChannel(ChannelId));

    /// <summary>Unscrambles one received WebSocket message and handles every RPC in it.</summary>
    public (IReadOnlyList<ScrChatEvent> Events, IReadOnlyList<byte[]> Replies) Receive(ReadOnlySpan<byte> wireMessage)
    {
        var events = new List<ScrChatEvent>();
        var replies = new List<byte[]>();
        foreach (var rpc in ClassicFrame.DecodeAll(ClassicEnvelope.Unscramble(wireMessage, _seed)))
        {
            if (rpc.Header.IsResponse)
            {
                var method = _pendingMethods.Remove(rpc.Header.Token, out var sent) ? sent : rpc.Header.Method;
                if (rpc.Header.Status != 0)
                {
                    events.Add(new ScrChatEvent.RequestFailed(method, rpc.Header.Token, rpc.Header.Status));
                }

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
            return new ScrChatEvent.Unhandled(rpc.Header.Service, rpc.Header.Method);
        }

        switch (rpc.Header.Method)
        {
            case LegacyChatService.CurrentChannelCallback:
                if (LegacyChatCallbacks.DecodeCurrentChannel(rpc.Body) is not { } channel)
                {
                    return null;
                }

                ChannelId = channel.Id;
                return new ScrChatEvent.ChannelEntered(channel);

            case LegacyChatService.ChannelListChangedCallback:
                return new ScrChatEvent.ChannelListChanged(LegacyChatCallbacks.DecodeChannelListChanges(rpc.Body));

            default:
                if (LegacyChatCallbacks.TryGetMessageKind(rpc.Header.Method, out var kind))
                {
                    return LegacyChatCallbacks.DecodeMessage(kind, rpc.Body) is { } message ? new ScrChatEvent.Message(message) : null;
                }

                return new ScrChatEvent.Unhandled(rpc.Header.Service, rpc.Header.Method);
        }
    }

    private byte[] Request((uint Method, byte[] Body) request)
    {
        _token = unchecked(_token + 1);
        _pendingMethods[_token] = request.Method;
        var header = new ClassicHeader(LegacyChatService.Hash, request.Method, _token, ClassicHeader.RequestRouting, 0, 0, false);
        return ClassicEnvelope.Scramble(ClassicFrame.Encode(header, request.Body), _seed);
    }
}
