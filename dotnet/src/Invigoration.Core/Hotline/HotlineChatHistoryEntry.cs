using System.Buffers.Binary;
using System.Text;

namespace Invigoration.Core.Hotline;

/// <summary>
/// One persisted chat message returned by the "2.5"-era chat history extension (Get Chat History,
/// transaction 700) — see
/// github.com/fogWraith/Hotline/blob/main/Docs/Protocol/Capabilities-Chat-History.md. Only ever
/// sent by a server that confirmed CAPABILITY_CHAT_HISTORY during login
/// (HotlineTransactionClient.SupportsChatHistory); a pre-2.5 (1.2.3+) server never sends these at
/// all, so this client's behavior there is completely unchanged.
/// </summary>
public sealed record HotlineChatHistoryEntry(
    ulong MessageId,
    DateTimeOffset Timestamp,
    bool IsAction,
    bool IsServerMessage,
    bool IsDeleted,
    ushort IconId,
    string Nickname,
    string Message)
{
    /// <summary>
    /// Parses one DATA_HISTORY_ENTRY's raw bytes: an 8-byte message ID, 8-byte signed Unix-seconds
    /// timestamp, 2-byte flags, 2-byte icon ID, 2-byte nick length, then the nick and (2-byte
    /// length-prefixed) message text — 22 bytes minimum before the variable fields. Any optional
    /// mini-TLV sub-fields that might follow the message body (none are defined yet in this
    /// version of the spec) are simply never read, which is itself the documented
    /// forward-compatible behavior — a parser that stops after the message body already ignores
    /// them correctly. Returns null (rather than throwing) for a truncated/malformed entry, so one
    /// bad entry doesn't take down an entire history batch.
    /// </summary>
    public static HotlineChatHistoryEntry? TryParse(byte[] data)
    {
        if (data.Length < 22)
        {
            return null;
        }

        var messageId = BinaryPrimitives.ReadUInt64BigEndian(data.AsSpan(0));
        var timestampSeconds = BinaryPrimitives.ReadInt64BigEndian(data.AsSpan(8));
        var flags = BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(16));
        var iconId = BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(18));
        var nickLength = BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(20));

        if (22 + nickLength + 2 > data.Length)
        {
            return null;
        }

        var nick = Encoding.UTF8.GetString(data, 22, nickLength);
        var msgLengthOffset = 22 + nickLength;
        var msgLength = BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(msgLengthOffset));

        if (24 + nickLength + msgLength > data.Length)
        {
            return null;
        }

        var message = Encoding.UTF8.GetString(data, 24 + nickLength, msgLength);

        DateTimeOffset timestamp;
        try
        {
            timestamp = DateTimeOffset.FromUnixTimeSeconds(timestampSeconds);
        }
        catch (ArgumentOutOfRangeException)
        {
            // A malformed/out-of-range timestamp shouldn't drop an otherwise-valid entry.
            timestamp = DateTimeOffset.UnixEpoch;
        }

        return new HotlineChatHistoryEntry(
            messageId,
            timestamp,
            IsAction: (flags & 0x0001) != 0,
            IsServerMessage: (flags & 0x0002) != 0,
            IsDeleted: (flags & 0x0004) != 0,
            iconId,
            nick,
            message);
    }
}
