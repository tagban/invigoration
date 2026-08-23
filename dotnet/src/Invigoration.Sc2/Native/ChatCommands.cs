using System.Buffers.Binary;
using System.Numerics;
using System.Text;
using Invigoration.Sc2.Chat;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

/// <summary>
/// Hand-rolled encoders for the native ("Sunken") chat records, ported from
/// core/src/native/protocol.rs. These are NOT schema-driven — the upstream
/// crate hand-writes every one of these builders directly against a
/// BitWriter, matching the "hand-rolled outbound" side of the codec. Every
/// method here was verified against golden hex vectors taken from that
/// crate's own unit tests (see ChatCommandsTests).
/// </summary>
public static class ChatCommands
{
    public const byte ChatSlot = 5;
    public const byte ToonSlot = 15;
    public const byte CacheSlot = 11;
    public const int ChannelIndexCount = 7;

    private const byte ChatJoinRequestCommand = 0;
    private const byte ChatLeaveRequestCommand = 2;
    private const byte ChatInviteAcceptCommand = 5;
    private const byte ChatInviteDeclineCommand = 6;
    private const byte ChatMessageCommand = 11;
    private const byte ChatWhisperSendCommand = 19;
    private const byte ToonSelectCommand = 5;
    private const byte CacheGetStreamItemsCommand = 9;

    private static BitWriter RecordWriter(byte command, byte serviceSlot)
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, command, serviceSlot);
        return writer;
    }

    public static byte[] ChatJoinPublic(ushort channelNameId, uint token, string locale)
    {
        if (locale.Length != 4)
        {
            throw new ArgumentException("Chat locale must be a four-character FourCC.", nameof(locale));
        }

        var writer = RecordWriter(ChatJoinRequestCommand, ChatSlot);
        writer.Write(2, 2);
        writer.Write(FourCc.Encode(locale), 32);
        writer.Write(channelNameId, 16);
        writer.Write(token, 32);
        writer.Align();
        return writer.ToBytes();
    }

    public static byte[] ChatJoinPrivate(string name, uint token)
    {
        if (name.Length == 0)
        {
            throw new ArgumentException("Private chat channel name cannot be empty.", nameof(name));
        }

        var writer = RecordWriter(ChatJoinRequestCommand, ChatSlot);
        writer.Write(0, 2);
        EncodeGeneratedUtf8(writer, name, lengthBits: 7, maximumBytes: 124, maximumCharacters: 31);
        writer.Write(token, 32);
        writer.Align();
        return writer.ToBytes();
    }

    public static byte[] ChatJoinClub(uint clubId, uint token)
    {
        var writer = RecordWriter(ChatJoinRequestCommand, ChatSlot);
        writer.Write(3, 2);
        writer.Write(0, 16);
        writer.Write(clubId, 32);
        writer.Write(token, 32);
        writer.Align();
        return writer.ToBytes();
    }

    public static byte[] ChatLeave(byte channelIndex)
    {
        RequireChannelIndex(channelIndex);
        var writer = RecordWriter(ChatLeaveRequestCommand, ChatSlot);
        writer.Write(channelIndex, 3);
        writer.Align();
        return writer.ToBytes();
    }

    public static byte[] ChatInviteAnswer(byte channelIndex, bool accept)
    {
        RequireChannelIndex(channelIndex);
        var command = accept ? ChatInviteAcceptCommand : ChatInviteDeclineCommand;
        var writer = RecordWriter(command, ChatSlot);
        writer.Write(channelIndex, 3);
        writer.Align();
        return writer.ToBytes();
    }

    public static byte[] ChatMessage(byte channelIndex, string body)
    {
        RequireChannelIndex(channelIndex);
        var writer = RecordWriter(ChatMessageCommand, ChatSlot);
        EncodeGeneratedUtf8(writer, body, lengthBits: 10, maximumBytes: 1020, maximumCharacters: 255);
        writer.Write(channelIndex, 3);
        writer.Align();
        return writer.ToBytes();
    }

    public static byte[] ChatWhisper(WhisperTarget target, string body)
    {
        var writer = RecordWriter(ChatWhisperSendCommand, ChatSlot);
        switch (target)
        {
            case WhisperTarget.Presence presence:
                writer.Write(0, 3);
                writer.Write(presence.PresenceId, 32);
                break;
            case WhisperTarget.Account account:
                writer.Write(3, 3);
                writer.Write(account.AccountId, 32);
                break;
            case WhisperTarget.ToonHandle handle:
                writer.Write(5, 3);
                writer.Write(handle.ProgramId, 32);
                writer.Write(handle.Region, 8);
                writer.Write(handle.Realm, 32);
                writer.Write(handle.Id, 64);
                break;
            case WhisperTarget.ToonName name:
                var nameBytes = Encoding.UTF8.GetBytes(name.Name);
                if (nameBytes.Length is < 2 or > 100)
                {
                    throw new ArgumentException("Whisper toon name must contain between 2 and 100 UTF-8 bytes.", nameof(target));
                }

                writer.Write(1, 3);
                writer.Write(name.Region, 8);
                writer.Write(name.ProgramId, 32);
                writer.Write(name.Realm, 32);
                writer.Write((ulong)(nameBytes.Length - 2), 7);
                writer.WriteBytes(nameBytes, aligned: true);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(target), "Unknown whisper target.");
        }

        EncodeGeneratedUtf8(writer, body, lengthBits: 10, maximumBytes: 1020, maximumCharacters: 255);
        writer.Align();
        return writer.ToBytes();
    }

    public static byte[] ToonSelect(string toonName, uint realm)
    {
        var rawName = Encoding.UTF8.GetBytes(toonName);
        var characterCount = toonName.EnumerateRunes().Count();
        if (characterCount is < 2 or > 25 || rawName.Length is < 2 or > 100)
        {
            throw new ArgumentException("Toon name must contain 2..=25 characters and 2..=100 UTF-8 bytes.", nameof(toonName));
        }

        var writer = RecordWriter(ToonSelectCommand, ToonSlot);
        writer.Write((ulong)(rawName.Length - 2), 7);
        writer.WriteBytes(rawName, aligned: true);
        WriteGeneratedChecksum(writer, width: 10, seed: 2);
        writer.Write(realm, 32);
        writer.Align();
        return writer.ToBytes();
    }

    /// <summary>
    /// Requests a Battle.net "cache stream" item — this is how the client pulls
    /// its bootstrap catalogs (Battle.net error strings, the public-channel
    /// list) at the start of ChatBootstrap, before selecting a toon. Ported
    /// from core/src/native/protocol.rs's cache_get_stream_items; upstream's
    /// own LiveChat::establish sends exactly two of these immediately on
    /// connecting (token 1: "BNET"/"ERRS"/"enUS" for the error catalog; token
    /// 2: "BNET"/"CONF"/"enUS" for the public-channel catalog, whose response
    /// is what resolves "General"'s numeric channel id). Response is decoded
    /// as a CacheStreamItems payload on the Cache slot/command 9.
    /// </summary>
    public static byte[] CacheGetStreamItems(uint token, string channel, string itemName, string locale)
    {
        if (channel.Length != 4 || itemName.Length != 4 || locale.Length != 4)
        {
            throw new ArgumentException("Cache stream channel, item name, and locale must each be a four-character FourCC.");
        }

        var writer = RecordWriter(CacheGetStreamItemsCommand, CacheSlot);
        writer.Write(token, 32);
        WriteGeneratedChecksum(writer, width: 23, seed: 7);
        writer.Write(0, 6);
        writer.Write(1, 1);
        writer.Write(FourCc.Encode(channel), 32);
        writer.Write(FourCc.Encode(itemName), 32);
        writer.Write(FourCc.Encode(locale), 32);
        writer.Write(0xFFFF_FFFF, 32);
        writer.Write(0, 1);
        writer.Align();
        return writer.ToBytes();
    }

    private static void RequireChannelIndex(byte channelIndex)
    {
        if (channelIndex >= ChannelIndexCount)
        {
            throw new ArgumentOutOfRangeException(nameof(channelIndex), "Chat channel index must be between 0 and 6.");
        }
    }

    /// <summary>
    /// Writes a bit-packed length prefix followed by the byte-aligned UTF-8 body — no minimum-byte
    /// offset is applied here (unlike the decoder side's <c>decode_generated_utf8</c>, which adds
    /// one back). Callers that need a minimum (e.g. toon names, min 2 bytes) write the biased
    /// length by hand instead of calling this helper, matching upstream.
    /// </summary>
    private static void EncodeGeneratedUtf8(BitWriter writer, string value, int lengthBits, int maximumBytes, int maximumCharacters)
    {
        var bytes = Encoding.UTF8.GetBytes(value);
        if (bytes.Length > maximumBytes || value.EnumerateRunes().Count() > maximumCharacters)
        {
            throw new ArgumentException("Generated native string exceeds its schema bound.", nameof(value));
        }

        writer.Write((ulong)bytes.Length, lengthBits);
        writer.WriteBytes(bytes, aligned: true);
    }

    /// <summary>
    /// Reproduces the upstream "generated checksum" trailer used after <c>toon_select</c>'s name
    /// field: reads back the four and two bytes immediately preceding the writer's current
    /// (byte-aligned) position, sums them with <paramref name="seed"/>, rotates left by 8 bits, and
    /// writes the low <paramref name="width"/> bits. Requires the writer to already be byte-aligned
    /// with at least 4 bytes written.
    /// </summary>
    private static void WriteGeneratedChecksum(BitWriter writer, int width, uint seed)
    {
        if (width is < 1 or > 32 || writer.Position < 32)
        {
            throw new InvalidOperationException("Generated checksum parameters are invalid.");
        }

        var byteIndex = writer.Position / 8;
        if (byteIndex < 4)
        {
            throw new InvalidOperationException("Generated checksum requires four preceding encoded bytes.");
        }

        var encoded = writer.ToBytes();
        var precedingFour = BinaryPrimitives.ReadUInt32LittleEndian(encoded.AsSpan(byteIndex - 4, 4));
        var precedingTwo = (uint)BinaryPrimitives.ReadUInt16LittleEndian(encoded.AsSpan(byteIndex - 2, 2));
        var checksum = unchecked(seed + precedingFour + precedingTwo);
        checksum = BitOperations.RotateLeft(checksum, 8);
        var mask = width == 32 ? uint.MaxValue : (1u << width) - 1;
        writer.Write(checksum & mask, width);
    }
}
