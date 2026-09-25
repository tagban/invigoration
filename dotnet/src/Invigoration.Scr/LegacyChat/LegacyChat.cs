using Invigoration.Sc2.Protobuf;
using Invigoration.Scr.Classic;

namespace Invigoration.Scr.LegacyChat;

/// <summary>
/// SC:R's chat service on the classic connection. Method IDs are 32-bit hashes,
/// sent as varints. See docs.bnet.cc's "StarCraft: Remastered chat".
/// </summary>
public static class LegacyChatService
{
    public const uint Hash = 0xF4E11A78;

    /// <summary>ConnectionService, whose method 3 is an echo the client must answer with the same body.</summary>
    public const uint ConnectionHash = 0x65446991;
    public const uint ConnectionEchoMethod = 3;

    public const uint SendMessageMethod = 0x5851FB2D;
    public const uint CommandMethod = 0x1FEE1493;
    public const uint ListChannelsMethod = 0x78D3F5A8;
    public const uint JoinListedChannelMethod = 0x00C47F9F;
    public const uint LeaveChannelMethod = 0x84F6DDA8;

    public const uint ChannelListChangedCallback = 0xC04DAC29;

    /// <summary>ForceJoinChannel: the server putting us in a channel, which confirms a join.</summary>
    public const uint CurrentChannelCallback = 0xC583300A;

    /// <summary>LeftChannel: the server taking us out of a channel. Body field 1 is its ID.</summary>
    public const uint LeftChannelCallback = 0xB07FD98A;

    public const uint SetOnlineMethod = 0xD5EBA117;
}

/// <summary>Which kind of text a LegacyChat callback carries.</summary>
public enum ScrMessageKind
{
    Whisper,
    Channel,
    Broadcast,
    Information,
    Error,
    Emote,
}

/// <summary>Request bodies for LegacyChat. Each returns the method to call and its body.</summary>
public static class LegacyChatRequests
{
    public static (uint Method, byte[] Body) SendMessage(ulong channelId, string text)
    {
        RequireText(text);
        var w = new ProtoWriter();
        w.WriteUInt64(1, channelId);
        w.WriteString(2, text);
        return (LegacyChatService.SendMessageMethod, w.ToArray());
    }

    /// <summary>A whisper. The whole text goes in one argument, never split on spaces.</summary>
    public static (uint Method, byte[] Body) Whisper(ulong channelId, string recipient, string text)
    {
        RequireText(recipient);
        RequireText(text);
        return Command(channelId, "whisper", recipient, text);
    }

    /// <summary>Joins a channel by name. The whole name is one argument.</summary>
    public static (uint Method, byte[] Body) JoinByName(ulong channelId, string channelName)
    {
        RequireText(channelName);
        return Command(channelId, "channel", channelName);
    }

    /// <summary>
    /// A typed slash command ("/whois KTBPA", "/kick name reason"), as the game sends it: the command
    /// word, then the first argument, then the rest of the line as one argument.
    /// </summary>
    public static (uint Method, byte[] Body) SlashCommand(ulong channelId, string text)
    {
        var parts = text.TrimStart('/').Split(' ', 3, StringSplitOptions.RemoveEmptyEntries);
        RequireText(parts.Length > 0 ? parts[0] : "");
        return Command(channelId, parts[0].ToLowerInvariant(), parts[1..]);
    }

    public static (uint Method, byte[] Body) ListChannels()
    {
        var w = new ProtoWriter();
        w.WriteUInt64(1, 0);
        return (LegacyChatService.ListChannelsMethod, w.ToArray());
    }

    public static (uint Method, byte[] Body) JoinListedChannel(ulong targetChannelId)
    {
        var w = new ProtoWriter();
        w.WriteUInt64(1, targetChannelId);
        return (LegacyChatService.JoinListedChannelMethod, w.ToArray());
    }

    public static (uint Method, byte[] Body) LeaveChannel(ulong channelId)
    {
        var w = new ProtoWriter();
        w.WriteUInt64(1, channelId);
        return (LegacyChatService.LeaveChannelMethod, w.ToArray());
    }

    private static (uint Method, byte[] Body) Command(ulong channelId, string command, params string[] arguments)
    {
        var w = new ProtoWriter();
        w.WriteUInt64(1, channelId);
        w.WriteString(2, command);
        foreach (var argument in arguments)
        {
            w.WriteString(3, argument);
        }

        return (LegacyChatService.CommandMethod, w.ToArray());
    }

    private static void RequireText(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            throw new ArgumentException("Text cannot be empty.", nameof(text));
        }
    }
}

/// <summary>A member of an SC:R channel.</summary>
public sealed record ScrMember(string Name, ulong Flags, IReadOnlyDictionary<string, string> Attributes, byte[]? Raw = null);

/// <summary>An SC:R channel as LegacyChat describes it.</summary>
public sealed record ScrChannel(ulong Id, string InternalName, string DisplayName, IReadOnlyList<ScrMember> Members);

/// <summary>One entry of a channel-list change: type 1 removes the channel, any other type adds or updates it.</summary>
public sealed record ScrChannelChange(ulong ChangeType, ScrChannel Channel)
{
    public bool IsRemoval => ChangeType == 1;
}

/// <summary>Text pulled out of a message callback. <see cref="Sender"/> is null when the callback held only one string.</summary>
public sealed record ScrMessage(ScrMessageKind Kind, string? Sender, string Text);

/// <summary>Decoders for LegacyChat's server calls, and the replies they need.</summary>
public static class LegacyChatCallbacks
{
    private static readonly Dictionary<uint, ScrMessageKind> MessageKinds = new()
    {
        [0xFA88D3E1] = ScrMessageKind.Whisper,
        [0x850B6EE3] = ScrMessageKind.Channel,
        [0xAA6957AA] = ScrMessageKind.Broadcast,
        [0x1580B7A1] = ScrMessageKind.Information,
        [0xD52809ED] = ScrMessageKind.Error,
        [0x632D6CFD] = ScrMessageKind.Emote,
    };

    public static bool TryGetMessageKind(uint method, out ScrMessageKind kind) => MessageKinds.TryGetValue(method, out kind);

    /// <summary>
    /// The reply every server call on the classic connection needs: same service,
    /// method, token, routing and object ID, marked as a response. The body is empty, except
    /// for ConnectionService's echo, which gets its own body back.
    /// </summary>
    public static byte[] Reply(ClassicRpc call)
    {
        var echo = call.Header.Service == LegacyChatService.ConnectionHash && call.Header.Method == LegacyChatService.ConnectionEchoMethod;
        // Same service, method, token, routing and object ID; marked as a response; no trace.
        var header = call.Header with { Routing = call.Header.Routing ?? ClassicHeader.RequestRouting, IsResponse = true, RequestTrace = null };
        return ClassicFrame.Encode(header, echo ? call.Body : []);
    }

    /// <summary>Field 1 of a body that carries just a channel ID, such as LeftChannel.</summary>
    public static ulong DecodeChannelId(byte[] body)
    {
        var r = new ProtoReader(body);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            if (field == 1 && type == WireType.Varint)
            {
                return r.ReadVarint();
            }

            r.Skip(type);
        }

        return 0;
    }

    /// <summary>Channel list changes (<see cref="LegacyChatService.ChannelListChangedCallback"/>).</summary>
    public static IReadOnlyList<ScrChannelChange> DecodeChannelListChanges(byte[] body)
    {
        var changes = new List<ScrChannelChange>();
        var r = new ProtoReader(body);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            if (field == 1 && type == WireType.LengthDelimited)
            {
                changes.Add(DecodeChange(r.ReadLengthDelimited()));
            }
            else
            {
                r.Skip(type);
            }
        }

        return changes;
    }

    /// <summary>The channel the client is in (<see cref="LegacyChatService.CurrentChannelCallback"/>), or null if the call didn't carry one.</summary>
    public static ScrChannel? DecodeCurrentChannel(byte[] body)
    {
        ScrChannel? channel = null;
        var r = new ProtoReader(body);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            if (field == 2 && type == WireType.LengthDelimited)
            {
                channel = DecodeChannel(r.ReadLengthDelimited());
            }
            else
            {
                r.Skip(type);
            }
        }

        return channel;
    }

    /// <summary>
    /// Pulls sender and text out of a message callback. The six kinds' exact
    /// layouts aren't mapped yet, so this reads by shape: every length-delimited
    /// field in order, up to four levels deep. A non-empty, valid UTF-8 string
    /// without control characters counts as text; anything else is tried as a
    /// nested message. The last text is the message and, if there's more than
    /// one, the first is the sender. Replace with typed decoders once captured.
    /// </summary>
    public static ScrMessage? DecodeMessage(ScrMessageKind kind, byte[] body)
    {
        var texts = new List<string>();
        CollectTexts(body, depth: 0, texts);

        // Server notices carry the recipient, us, where a talk or whisper carries its sender, so
        // they get no sender at all.
        var hasSender = kind is ScrMessageKind.Channel or ScrMessageKind.Whisper or ScrMessageKind.Emote;
        return texts.Count switch
        {
            0 => null,
            1 => new ScrMessage(kind, null, texts[0]),
            _ => new ScrMessage(kind, hasSender ? texts[0] : null, texts[^1]),
        };
    }

    private static void CollectTexts(byte[] message, int depth, List<string> texts)
    {
        if (depth >= 4)
        {
            return;
        }

        var r = new ProtoReader(message);
        try
        {
            while (r.HasMore)
            {
                var (_, type) = r.ReadTag();
                if (type != WireType.LengthDelimited)
                {
                    r.Skip(type);
                    continue;
                }

                var bytes = r.ReadLengthDelimited();
                if (TryReadText(bytes, out var text))
                {
                    texts.Add(text);
                }
                else
                {
                    CollectTexts(bytes, depth + 1, texts);
                }
            }
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or IndexOutOfRangeException)
        {
            // Not a protobuf message after all: whatever was collected before this point stands.
        }
    }

    private static readonly System.Text.UTF8Encoding StrictUtf8 = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

    private static bool TryReadText(byte[] bytes, out string text)
    {
        text = "";
        if (bytes.Length == 0)
        {
            return false;
        }

        try
        {
            text = StrictUtf8.GetString(bytes);
        }
        catch (System.Text.DecoderFallbackException)
        {
            return false;
        }

        return !text.Any(char.IsControl);
    }

    private static ScrChannelChange DecodeChange(byte[] data)
    {
        ulong changeType = 0;
        ScrChannel? channel = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: changeType = r.ReadVarint(); break;
                case 2 when type == WireType.LengthDelimited: channel = DecodeChannel(r.ReadLengthDelimited()); break;
                default: r.Skip(type); break;
            }
        }

        return new ScrChannelChange(changeType, channel ?? new ScrChannel(0, "", "", []));
    }

    private static ScrChannel DecodeChannel(byte[] data)
    {
        ulong id = 0;
        string internalName = "", displayName = "";
        var members = new List<ScrMember>();
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: id = r.ReadVarint(); break;
                case 2: internalName = r.ReadString(); break;
                case 3: members.Add(DecodeMember(r.ReadLengthDelimited())); break;
                case 5: displayName = r.ReadString(); break;
                default: r.Skip(type); break;
            }
        }

        // The channel list often carries only the internal name; it doubles as the display name then.
        return new ScrChannel(id, internalName, displayName.Length > 0 ? displayName : internalName, members);
    }

    private static ScrMember DecodeMember(byte[] data)
    {
        var name = "";
        ulong flags = 0;
        var attributes = new Dictionary<string, string>();
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: name = r.ReadString(); break;
                case 2: flags = r.ReadVarint(); break;
                case 3:
                    var (key, value) = DecodeAttribute(r.ReadLengthDelimited());
                    attributes[key] = value;
                    break;
                default: r.Skip(type); break;
            }
        }

        return new ScrMember(name, flags, attributes, data);
    }

    private static (string Name, string Value) DecodeAttribute(byte[] data)
    {
        string name = "", value = "";
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: name = r.ReadString(); break;
                case 2: value = r.ReadString(); break;
                default: r.Skip(type); break;
            }
        }

        return (name, value);
    }
}
