using Invigoration.Sc2.Front;

namespace Invigoration.Diablo;

/// <summary>Something that happened in D2R or D4 chat.</summary>
public abstract record DiabloChatEvent
{
    private DiabloChatEvent()
    {
    }

    public sealed record Message(string SenderName, DiabloChatMessage Value) : DiabloChatEvent;

    public sealed record MemberJoined(DiabloMember Member) : DiabloChatEvent;

    public sealed record MemberLeft(DiabloMember Member) : DiabloChatEvent;

    public sealed record Whisper(DiabloWhisper Value) : DiabloChatEvent;

    public sealed record ChannelTypesListed(IReadOnlyList<DiabloChannelType> Types) : DiabloChatEvent;

    /// <summary>A subscribe succeeded: this is now the current channel.</summary>
    public sealed record ChannelJoined(DiabloChannelDescription Channel) : DiabloChatEvent;

    public sealed record RequestFailed(string What, uint Status) : DiabloChatEvent;

    /// <summary>The server asked us to disconnect.</summary>
    public sealed record DisconnectRequested : DiabloChatEvent;
}

/// <summary>
/// D2R or D4 chat on an already-open, signed-in Front WebSocket. Every method
/// returns the binary WebSocket messages to send, in order; it owns no socket.
/// Joining a channel is Find, then (once both the Find reply and the channel's
/// description have arrived, in either order) Subscribe. A whisper is a
/// BattleTag lookup, then the whisper itself. Current channel and roster only
/// change once a subscribe succeeds.
/// </summary>
public sealed class DiabloChatSession
{
    private readonly DiabloChatRequests _requests;
    private readonly Dictionary<uint, Pending> _pending = new();
    private readonly Dictionary<GameAccount, DiabloMember> _roster = new();
    private uint _token;
    private JoinAttempt? _join;

    private abstract record Pending(string What)
    {
        public sealed record Plain(string What) : Pending(What);
        public sealed record ListTypes() : Pending("channel list");
        public sealed record Find() : Pending("find channel");
        public sealed record Subscribe(DiabloChannelDescription Description) : Pending("join channel");
        public sealed record Resolve(string BattleTag, string Text) : Pending("look up " + BattleTag);
    }

    private sealed class JoinAttempt(UniqueChannelType type, string identity)
    {
        public UniqueChannelType Type { get; } = type;
        public string Identity { get; } = identity;
        public bool FindAnswered { get; set; }
        public DiabloChannelDescription? Description { get; set; }
    }

    /// <param name="token">The Front connection's request counter as sign-in left it. Each request increments it first.</param>
    public DiabloChatSession(DiabloGame game, GameAccount account, DiabloChannelId channel, uint token, IEnumerable<DiabloMember>? members = null)
    {
        _requests = new DiabloChatRequests(game, account);
        Channel = channel;
        _token = token;
        foreach (var member in members ?? [])
        {
            if (member.Handle is { } handle)
            {
                _roster[handle] = member;
            }
        }
    }

    public DiabloChannelId Channel { get; private set; }

    public IReadOnlyCollection<DiabloMember> Members => _roster.Values;

    /// <summary>How often to send <see cref="KeepAlive"/>.</summary>
    public static readonly TimeSpan KeepAliveInterval = TimeSpan.FromSeconds(30);

    public byte[] SendMessage(string text) =>
        Call(DiabloServices.ChannelService, DiabloServices.SendMessageMethod, _requests.SendMessage(Channel, text), new Pending.Plain("send message"));

    public byte[] ListChannelTypes() =>
        Call(DiabloServices.ChannelService, DiabloServices.ListChannelTypesMethod, _requests.ListChannelTypes(), new Pending.ListTypes());

    /// <summary>Starts joining a public channel by the identity <see cref="ListChannelTypes"/> returned.</summary>
    public byte[] JoinChannel(UniqueChannelType type, string identity)
    {
        _join = new JoinAttempt(type, identity);
        return Call(DiabloServices.ChannelService, DiabloServices.FindChannelMethod, _requests.FindChannel(type, identity), new Pending.Find());
    }

    /// <summary>Starts a whisper: looks the BattleTag up, then sends once its account ID is known.</summary>
    public byte[] Whisper(string battleTag, string text) =>
        Call(DiabloServices.AccountService, DiabloServices.ResolveBattleTagMethod, DiabloChatRequests.ResolveBattleTag(battleTag), new Pending.Resolve(battleTag, text));

    /// <summary>ConnectionService 5, every <see cref="KeepAliveInterval"/>. No reply is expected.</summary>
    public byte[] KeepAlive()
    {
        _token = unchecked(_token + 1);
        return FrontFrame.Encode(RequestHeader(DiabloServices.ConnectionService, DiabloServices.KeepAliveMethod, 0), []);
    }

    /// <summary>Handles one received WebSocket message.</summary>
    public (IReadOnlyList<DiabloChatEvent> Events, IReadOnlyList<byte[]> Outgoing) Receive(byte[] message)
    {
        var events = new List<DiabloChatEvent>();
        var outgoing = new List<byte[]>();
        var (header, body) = FrontFrame.Decode(message);
        if (header.ServiceId == 254 || header.IsResponse == true)
        {
            HandleReply(header, body, events, outgoing);
        }
        else
        {
            HandleCall(header, body, events, outgoing);
        }

        return (events, outgoing);
    }

    private void HandleReply(Header header, byte[] body, List<DiabloChatEvent> events, List<byte[]> outgoing)
    {
        if (!_pending.Remove(header.Token, out var pending))
        {
            return;
        }

        if (header.Status is { } status and not 0)
        {
            if (pending is Pending.Find or Pending.Subscribe)
            {
                _join = null;
            }

            events.Add(new DiabloChatEvent.RequestFailed(pending.What, status));
            return;
        }

        switch (pending)
        {
            case Pending.ListTypes:
                events.Add(new DiabloChatEvent.ChannelTypesListed(DiabloChatDecoders.DecodeChannelTypes(body)));
                break;
            case Pending.Find when _join is not null:
                _join.FindAnswered = true;
                TrySubscribe(outgoing);
                break;
            case Pending.Subscribe subscribe:
                var joined = DiabloChatDecoders.DecodeSubscribeReply(body) ?? subscribe.Description;
                if (subscribe.Description.Channel is { } channel)
                {
                    Channel = channel;
                }

                _roster.Clear();
                foreach (var member in joined.Members.Count > 0 ? joined.Members : subscribe.Description.Members)
                {
                    if (member.Handle is { } handle)
                    {
                        _roster[handle] = member;
                    }
                }

                _join = null;
                events.Add(new DiabloChatEvent.ChannelJoined(joined));
                break;
            case Pending.Resolve resolve:
                if (DiabloChatDecoders.DecodeResolvedAccountId(body) is { } accountId)
                {
                    outgoing.Add(Call(DiabloServices.WhisperService, DiabloServices.SendWhisperMethod,
                        DiabloChatRequests.Whisper(accountId, resolve.Text), new Pending.Plain("whisper " + resolve.BattleTag)));
                }
                else
                {
                    events.Add(new DiabloChatEvent.RequestFailed(resolve.What, 0));
                }

                break;
        }
    }

    private void HandleCall(Header header, byte[] body, List<DiabloChatEvent> events, List<byte[]> outgoing)
    {
        var service = header.ServiceHash;
        var method = header.MethodId;
        if (service == DiabloServices.ConnectionService)
        {
            if (method == DiabloServices.EchoMethod)
            {
                outgoing.Add(FrontFrame.Encode(new Header { ServiceId = 254, Token = header.Token, IsResponse = true }, DiabloChatRequests.EchoReply(body)));
            }
            else if (method == DiabloServices.DisconnectRequestMethod)
            {
                events.Add(new DiabloChatEvent.DisconnectRequested());
            }

            return;
        }

        if (service == DiabloServices.ChannelListener)
        {
            HandleChannelCall(method, body, events);
        }
        else if (service == DiabloServices.MembershipListener && method == DiabloServices.ChannelDescriptionCallback && _join is not null)
        {
            if (DiabloChatDecoders.DecodeMembershipDescription(body) is { } description && description.Identity == _join.Identity)
            {
                _join.Description = description;
                TrySubscribe(outgoing);
            }
        }
        else if (service == DiabloServices.WhisperListener && method == DiabloServices.WhisperCallback)
        {
            if (DiabloChatDecoders.DecodeWhisper(body) is { } whisper)
            {
                events.Add(new DiabloChatEvent.Whisper(whisper));
            }
        }
    }

    private void HandleChannelCall(uint? method, byte[] body, List<DiabloChatEvent> events)
    {
        switch (method)
        {
            case DiabloServices.MessageCallback:
                if (DiabloChatDecoders.DecodeMessage(body) is { } message && IsCurrent(message.Channel))
                {
                    var name = message.Sender is { } sender && _roster.TryGetValue(sender, out var member) ? member.Name : "";
                    events.Add(new DiabloChatEvent.Message(name, message));
                }

                break;
            case DiabloServices.MemberAddedCallback:
                var (addedIn, added) = DiabloChatDecoders.DecodeMemberAdded(body);
                if (added?.Handle is { } addedHandle && IsCurrent(addedIn))
                {
                    var isNew = !_roster.ContainsKey(addedHandle);
                    _roster[addedHandle] = added;
                    if (isNew)
                    {
                        events.Add(new DiabloChatEvent.MemberJoined(added));
                    }
                }

                break;
            case DiabloServices.MemberRemovedCallback:
                var (removedFrom, removed) = DiabloChatDecoders.DecodeMemberRemoved(body);
                if (removed is { } removedHandle && IsCurrent(removedFrom) && _roster.Remove(removedHandle, out var left))
                {
                    events.Add(new DiabloChatEvent.MemberLeft(left));
                }

                break;
        }
    }

    /// <summary>Subscribe once both halves of a Find have arrived.</summary>
    private void TrySubscribe(List<byte[]> outgoing)
    {
        if (_join is not { FindAnswered: true, Description: { Channel: { } channel } description })
        {
            return;
        }

        outgoing.Add(Call(DiabloServices.ChannelService, DiabloServices.SubscribeMethod, _requests.ChannelAction(channel), new Pending.Subscribe(description)));
        _join.FindAnswered = false;
    }

    private bool IsCurrent(DiabloChannelId? channel) => channel is null || channel == Channel;

    private byte[] Call(uint service, uint method, byte[] body, Pending pending)
    {
        _token = unchecked(_token + 1);
        _pending[_token] = pending;
        return FrontFrame.Encode(RequestHeader(service, method, (uint)body.Length), body);
    }

    private Header RequestHeader(uint service, uint method, uint size) =>
        new() { ServiceId = 0, MethodId = method, Token = _token, Size = size, ServiceHash = service };
}
