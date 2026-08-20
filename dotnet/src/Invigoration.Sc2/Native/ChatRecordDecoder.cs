using System.Text;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

/// <summary>
/// Hand-rolled decoders for the native ("Sunken") chat records that were
/// portable without upstream's ~500KB schema table, ported from
/// core/src/native/decode.rs. Covers ChatInviteNotify (4), JoinNotify2 (27),
/// MessageRecv (11), and WhisperRecv/WhisperEchoRecv (19/30) — either fully
/// hand-rolled upstream or "constant-foldable" (their reflective field
/// lookups resolve, per the static wire schema, to fixed bit widths with
/// zero-bit empty base structs). MembershipChangeNotify (command 1) is
/// decoded separately by <see cref="MembershipChangeDecoder"/>, once the
/// field-level schema for its previously-blocked status variants became
/// available via the SC2Docs type registry.
/// </summary>
public static class ChatRecordDecoder
{
    public static ChatInviteRecord DecodeChatInvite(BitReader reader)
    {
        var channelType = (byte)reader.Read(4);
        var inviterPresence = (uint)reader.Read(32);
        var channelIndex = (byte)reader.Read(3);
        return new ChatInviteRecord(channelType, inviterPresence, channelIndex);
    }

    public static ChatJoinRecord DecodeChatJoin(BitReader reader)
    {
        var isFailure = reader.Read(1) != 0;
        if (!isFailure)
        {
            var memberHandle = (uint)reader.Read(32);
            var channelIndex = (byte)reader.Read(3);
            if (channelIndex >= ChatCommands.ChannelIndexCount)
            {
                throw new InvalidOperationException("Chat join channel index is invalid.");
            }

            reader.Read(32); // conference_id, discarded
            reader.Read(32); // owner_id, discarded
            var channelType = (byte)reader.Read(4);

            ushort? channelNameId = null;
            ushort? channelShardIndex = null;
            if (reader.Read(1) != 0)
            {
                var (nameId, shardIndex) = DecodeChannelName(reader);
                channelNameId = nameId;
                channelShardIndex = shardIndex;
            }

            if (reader.Read(1) != 0)
            {
                DecodeChannelConfig(reader);
            }

            if (reader.Read(1) != 0)
            {
                reader.Read(32); // reserved
            }

            uint? token = reader.Read(1) != 0 ? (uint)reader.Read(32) : null;

            return new ChatJoinRecord(true, channelIndex, memberHandle, channelType, channelNameId, channelShardIndex, null, token);
        }
        else
        {
            var reason = (ushort)reader.Read(16);
            byte? channelType = reader.Read(1) == 0 ? null : (byte)reader.Read(4);
            uint? token = reader.Read(1) != 0 ? (uint)reader.Read(32) : null;
            return new ChatJoinRecord(false, null, null, channelType, null, null, reason, token);
        }
    }

    public static ChatMessageRecord DecodeChatMessage(BitReader reader)
    {
        var memberHandle = (uint)reader.Read(32);
        var body = DecodeGeneratedUtf8(reader, lengthBits: 10, minimumBytes: 0, maximumBytes: 1020, maximumCharacters: 255);
        var channelIndex = (byte)reader.Read(3);
        return new ChatMessageRecord(memberHandle, body, channelIndex);
    }

    /// <summary>
    /// Decodes WhisperRecv/WhisperEchoRecv. Upstream's decoder hard-codes the
    /// base-class field name "Whisper" even for the echo route, which — per
    /// the schema — should fail (WhisperEchoRecv's base field is actually
    /// named "WhisperEcho"). Both base structs are empty (zero bits on the
    /// wire), so this decoder sidesteps the bug entirely by not looking the
    /// base field up at all; it works correctly for both routes.
    /// </summary>
    public static ChatWhisperRecord DecodeChatWhisper(BitReader reader)
    {
        var region = (byte)reader.Read(8);
        var programId = (uint)reader.Read(32);
        var realm = (uint)reader.Read(32);
        var name = DecodeGeneratedUtf8(reader, lengthBits: 7, minimumBytes: 2, maximumBytes: 100, maximumCharacters: 25);
        var body = DecodeGeneratedUtf8(reader, lengthBits: 10, minimumBytes: 0, maximumBytes: 1020, maximumCharacters: 255);
        return new ChatWhisperRecord(region, programId, realm, name, body);
    }

    private static (ushort? PublicChannelNameId, ushort ShardIndex) DecodeChannelName(BitReader reader)
    {
        var shardIndex = (ushort)reader.Read(16);
        reader.Read(29);
        var selector = reader.Read(2);
        ushort? publicId;
        switch (selector)
        {
            case 2:
                reader.Read(32); // locale fourcc, discarded
                publicId = (ushort)reader.Read(16);
                break;
            case 1:
            case 3:
                reader.Read(16);
                reader.Read(32);
                publicId = null;
                break;
            case 0:
                DecodeGeneratedUtf8(reader, lengthBits: 7, minimumBytes: 0, maximumBytes: 124, maximumCharacters: 31);
                publicId = null;
                break;
            default:
                throw new InvalidOperationException("Two bits have only four values.");
        }

        return (publicId, shardIndex);
    }

    private static void DecodeChannelConfig(BitReader reader)
    {
        reader.Read(8);
        reader.Read(16);
        var programs = (int)reader.Read(3);
        if (programs > 4)
        {
            throw new InvalidOperationException("Channel config has too many programs.");
        }

        for (var i = 0; i < programs; i++)
        {
            reader.Read(32);
        }

        reader.Read(32);
        var realms = (int)reader.Read(3);
        if (realms > 4)
        {
            throw new InvalidOperationException("Channel config has too many realms.");
        }

        for (var i = 0; i < realms; i++)
        {
            reader.Read(32);
        }

        reader.Read(32);
    }

    private static string DecodeGeneratedUtf8(BitReader reader, int lengthBits, int minimumBytes, int maximumBytes, int maximumCharacters)
    {
        var byteCount = (int)reader.Read(lengthBits) + minimumBytes;
        if (byteCount > maximumBytes)
        {
            throw new InvalidOperationException("Generated native string is too long.");
        }

        var bytes = reader.ReadBytes(byteCount, aligned: true);
        var value = Encoding.UTF8.GetString(bytes);
        if (value.EnumerateRunes().Count() > maximumCharacters)
        {
            throw new InvalidOperationException("Generated native string has too many characters.");
        }

        return value;
    }
}
