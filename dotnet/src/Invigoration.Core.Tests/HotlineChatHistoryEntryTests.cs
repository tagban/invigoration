using System.Buffers.Binary;
using System.Text;
using Invigoration.Core.Hotline;

namespace Invigoration.Core.Tests;

public class HotlineChatHistoryEntryTests
{
    private static byte[] BuildEntry(ulong messageId, long timestampSeconds, ushort flags, ushort iconId, string nick, string message, byte[]? trailingJunk = null)
    {
        var nickBytes = Encoding.UTF8.GetBytes(nick);
        var msgBytes = Encoding.UTF8.GetBytes(message);
        var buffer = new List<byte>();

        Span<byte> eight = stackalloc byte[8];
        BinaryPrimitives.WriteUInt64BigEndian(eight, messageId);
        buffer.AddRange(eight.ToArray());

        BinaryPrimitives.WriteInt64BigEndian(eight, timestampSeconds);
        buffer.AddRange(eight.ToArray());

        Span<byte> two = stackalloc byte[2];
        BinaryPrimitives.WriteUInt16BigEndian(two, flags);
        buffer.AddRange(two.ToArray());
        BinaryPrimitives.WriteUInt16BigEndian(two, iconId);
        buffer.AddRange(two.ToArray());
        BinaryPrimitives.WriteUInt16BigEndian(two, (ushort)nickBytes.Length);
        buffer.AddRange(two.ToArray());

        buffer.AddRange(nickBytes);
        BinaryPrimitives.WriteUInt16BigEndian(two, (ushort)msgBytes.Length);
        buffer.AddRange(two.ToArray());
        buffer.AddRange(msgBytes);

        if (trailingJunk is not null)
        {
            buffer.AddRange(trailingJunk);
        }

        return [.. buffer];
    }

    [Fact]
    public void TryParse_WellFormedEntry_RoundTripsAllFields()
    {
        var data = BuildEntry(42, 1700000000, flags: 0, iconId: 414, "Tagban", "hello world");

        var entry = HotlineChatHistoryEntry.TryParse(data);

        Assert.NotNull(entry);
        Assert.Equal(42UL, entry!.MessageId);
        Assert.Equal(DateTimeOffset.FromUnixTimeSeconds(1700000000), entry.Timestamp);
        Assert.False(entry.IsAction);
        Assert.False(entry.IsServerMessage);
        Assert.False(entry.IsDeleted);
        Assert.Equal((ushort)414, entry.IconId);
        Assert.Equal("Tagban", entry.Nickname);
        Assert.Equal("hello world", entry.Message);
    }

    [Fact]
    public void TryParse_ActionFlagSet_IsActionTrue()
    {
        var data = BuildEntry(1, 1700000000, flags: 0x0001, iconId: 0, "Tagban", "waves");
        Assert.True(HotlineChatHistoryEntry.TryParse(data)!.IsAction);
    }

    [Fact]
    public void TryParse_ServerMessageFlagSet_IsServerMessageTrue()
    {
        var data = BuildEntry(1, 1700000000, flags: 0x0002, iconId: 0, "", "Server restarting soon");
        var entry = HotlineChatHistoryEntry.TryParse(data);
        Assert.True(entry!.IsServerMessage);
        Assert.Equal("", entry.Nickname);
    }

    [Fact]
    public void TryParse_DeletedFlagSet_IsDeletedTrue()
    {
        var data = BuildEntry(1, 1700000000, flags: 0x0004, iconId: 0, "", "");
        Assert.True(HotlineChatHistoryEntry.TryParse(data)!.IsDeleted);
    }

    [Fact]
    public void TryParse_IgnoresTrailingMiniTlvSubFields()
    {
        // Forward-compatibility: a future sub-field this client doesn't understand must not break
        // parsing of the fixed/variable fields that come before it.
        var data = BuildEntry(1, 1700000000, flags: 0, iconId: 0, "Tagban", "hi", trailingJunk: [0x00, 0x01, 0x00, 0x02, 0xAB, 0xCD]);
        var entry = HotlineChatHistoryEntry.TryParse(data);
        Assert.NotNull(entry);
        Assert.Equal("hi", entry!.Message);
    }

    [Fact]
    public void TryParse_TooShortForFixedHeader_ReturnsNull()
    {
        Assert.Null(HotlineChatHistoryEntry.TryParse(new byte[10]));
    }

    [Fact]
    public void TryParse_TruncatedMessageBody_ReturnsNull()
    {
        var full = BuildEntry(1, 1700000000, flags: 0, iconId: 0, "Tagban", "hello world");
        var truncated = full[..(full.Length - 5)];
        Assert.Null(HotlineChatHistoryEntry.TryParse(truncated));
    }
}
