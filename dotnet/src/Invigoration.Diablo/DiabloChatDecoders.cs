namespace Invigoration.Diablo;

/// <summary>A channel member. <see cref="Handle"/> identifies them in messages; match on all three of its parts.</summary>
public sealed record DiabloMember(GameAccount? Handle, string Name, ulong? AccountId);

public sealed record DiabloChatMessage(DiabloChannelId? Channel, GameAccount? Sender, string Text);

public sealed record DiabloWhisper(ulong? SenderAccountId, string? SenderBattleTag, string Text);

public sealed record DiabloChannelType(UniqueChannelType? Type, string Name, string Identity);

public sealed record DiabloChannelDescription(DiabloChannelId? Channel, UniqueChannelType? Type, string Name, string? Identity, IReadOnlyList<DiabloMember> Members);

/// <summary>
/// Decoders for what the server sends D2R and D4 chat clients. Field numbers
/// are documented; where D2R, D4 or older replies differ, every known variant
/// is accepted. See docs.bnet.cc's "Diablo II: Resurrected and Diablo IV chat".
/// </summary>
public static class DiabloChatDecoders
{
    /// <summary>ChannelListener 10: a channel message.</summary>
    public static DiabloChatMessage? DecodeMessage(byte[] body)
    {
        var fields = ProtoFields.Parse(body);
        if (fields.Message(4) is not { } message || ReadText(message, 3, 4) is not { } text)
        {
            return null;
        }

        return new DiabloChatMessage(ChannelOf(fields), message.Message(1) is { } sender ? DecodeHandle(sender) : null, text);
    }

    /// <summary>ChannelListener 3: a member joined or changed.</summary>
    public static (DiabloChannelId? Channel, DiabloMember? Member) DecodeMemberAdded(byte[] body)
    {
        var fields = ProtoFields.Parse(body);
        return (ChannelOf(fields), fields.Message(4) is { } member ? DecodeMember(member) : null);
    }

    /// <summary>ChannelListener 4: a member left. Returns their handle.</summary>
    public static (DiabloChannelId? Channel, GameAccount? Handle) DecodeMemberRemoved(byte[] body)
    {
        var fields = ProtoFields.Parse(body);
        return (ChannelOf(fields), fields.Message(4) is { } handle ? DecodeHandle(handle) : null);
    }

    /// <summary>WhisperListener 1.</summary>
    public static DiabloWhisper? DecodeWhisper(byte[] body)
    {
        var fields = ProtoFields.Parse(body);
        if (fields.Message(2) is not { } whisper || ReadText(whisper, 5, 3) is not { } text)
        {
            return null;
        }

        return new DiabloWhisper(whisper.Number(2), fields.String(3), text);
    }

    /// <summary>The reply to AccountService 13: the account ID in field 12, nested in field 1 or, in older replies, directly.</summary>
    public static ulong? DecodeResolvedAccountId(byte[] body)
    {
        var fields = ProtoFields.Parse(body);
        return fields.Message(12) is { } nested ? nested.Number(1) : fields.Number(12);
    }

    /// <summary>The reply to ChannelService 5: one entry per public channel.</summary>
    public static IReadOnlyList<DiabloChannelType> DecodeChannelTypes(byte[] body) =>
        ProtoFields.Parse(body).Messages(1)
            .Select(t => new DiabloChannelType(t.Message(1) is { } u ? DecodeUniqueType(u) : null, t.String(2) ?? "", t.String(3) ?? ""))
            .ToList();

    /// <summary>ChannelMembershipListener 1: a channel's description, in field 3.</summary>
    public static DiabloChannelDescription? DecodeMembershipDescription(byte[] body) =>
        ProtoFields.Parse(body).Message(3) is { } description ? DecodeDescription(description) : null;

    /// <summary>
    /// The reply to ChannelService 10 (subscribe): the channel and its roster in
    /// field 1. Assumed to share the membership description's layout, which
    /// isn't confirmed yet; check it against a live join.
    /// </summary>
    public static DiabloChannelDescription? DecodeSubscribeReply(byte[] body) =>
        ProtoFields.Parse(body).Message(1) is { } snapshot ? DecodeDescription(snapshot) : null;

    public static DiabloChannelId DecodeChannelId(ProtoFields fields)
    {
        var host = fields.Message(2);
        return new DiabloChannelId(host?.Number(1) ?? 0, host?.Number(2) ?? 0, (uint)(fields.Number(3) ?? 0), (uint)(fields.Number(4) ?? 0));
    }

    /// <summary>A handle, whether D2R's fixed-width or D4's varint encoding.</summary>
    public static GameAccount DecodeHandle(ProtoFields fields) =>
        new(fields.Number(1) ?? 0, (uint)(fields.Number(2) ?? 0), (uint)(fields.Number(3) ?? 0));

    private static DiabloChannelDescription DecodeDescription(ProtoFields d) =>
        new(
            d.Message(1) is { } channel ? DecodeChannelId(channel) : null,
            d.Message(2) is { } type ? DecodeUniqueType(type) : null,
            d.String(3) ?? "",
            d.Message(110)?.String(1),
            d.Messages(6).Select(DecodeMember).ToList());

    private static UniqueChannelType DecodeUniqueType(ProtoFields fields) =>
        new((uint)(fields.Number(2) ?? 0), fields.String(3) ?? "");

    /// <summary>D2R usually puts the member's handle in field 1, D4 in field 7 with field 1 as a fallback.</summary>
    private static DiabloMember DecodeMember(ProtoFields member)
    {
        var handle = member.Message(7) ?? member.Message(1);
        return new DiabloMember(handle is null ? null : DecodeHandle(handle), member.String(2) ?? "", member.Number(6));
    }

    private static DiabloChannelId? ChannelOf(ProtoFields fields) =>
        fields.Message(3) is { } channel ? DecodeChannelId(channel) : null;

    /// <summary>
    /// Text in the first of <paramref name="numbers"/> that holds some. It may be
    /// the string itself or, mirroring how it's sent, a message whose field 4 is
    /// the string; which one the server uses isn't confirmed, so both are read.
    /// </summary>
    private static string? ReadText(ProtoFields fields, params int[] numbers)
    {
        foreach (var number in numbers)
        {
            if (fields.Bytes(number) is not { } bytes)
            {
                continue;
            }

            if (TryNested(bytes) is { } nested)
            {
                return nested;
            }

            return System.Text.Encoding.UTF8.GetString(bytes);
        }

        return null;
    }

    private static string? TryNested(byte[] bytes)
    {
        try
        {
            var inner = ProtoFields.Parse(bytes);
            return inner.String(4) is { } text && !text.Any(char.IsControl) ? text : null;
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or IndexOutOfRangeException)
        {
            return null;
        }
    }
}
