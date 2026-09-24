using System.Text;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

/// <summary>
/// One entry in a ToonBlockNotify snapshot (Battlenet::Friends::ToonBlockContainer) —
/// a toon being added to or removed from the account's block list.
/// </summary>
public sealed record ToonBlockEntry(ToonFullName Toon, bool IsRemove);

/// <summary>
/// Decoded payload of ToonBlockNotify (Friends slot, command 33) — the
/// account's toon-block-list snapshot/update, sent unprompted during
/// ChatBootstrap alongside FriendsList/ToonsOfFriends. A 7-bit entry count
/// (capped at 64), then per entry the 1-bit Add/Remove choice FIRST, then a
/// Battlenet::Toon::FullName (region, program, realm, then a 7-bit name
/// length biased +2, byte-aligned). This decoder used to read the name with
/// a 5-bit length and the choice bit after it, which turned the one real
/// capture into program 10650 and the name "\x02TrumpFl". The order and width
/// here read the same bytes as region 1, program "S2", realm 1 and a 16-byte
/// name starting "TrumpFlat"; the capture stops partway through that name.
/// The 7-bit width is the one the real ToonsOfFriends capture confirms for
/// the same Battlenet::Toon::FullName type.
/// </summary>
public sealed record ToonBlockNotifyRecord(IReadOnlyList<ToonBlockEntry> Entries, bool? Complete);

/// <summary>
/// Hand-rolled decoders for the native ("Sunken") Friends records
/// (FriendsListNotify5 — Friends slot, command 30 — and ToonsOfFriendsNotify
/// — Friends slot, command 6), ported from core/src/native/decode.rs's
/// friends_list_with_provenance/trace_friend_container/trace_friend_account/
/// trace_friend_character/trace_friend_custom_message and
/// friend_toons_with_provenance. Same "schema-free hand trace" approach as
/// <see cref="ChatRecordDecoder"/> and <see cref="MembershipChangeDecoder"/> —
/// see <see cref="RecordStream"/>'s remarks on why this is necessary at all.
///
/// FriendContainer5::Account's optional m_fullName (an AccountFullName) is
/// two strings, given name then surname, each an 8-bit byte count followed
/// by byte-aligned UTF-8. It used to be treated as unknowable without SC2's
/// embedded schema; <see cref="DecodeAccountFullName"/> now reads it.
/// </summary>
public static class FriendsRecordDecoder
{
    public static ToonBlockNotifyRecord DecodeToonBlockNotify(BitReader reader)
    {
        var count = (int)reader.Read(7);
        if (count > 64)
        {
            throw new InvalidOperationException("Toon block snapshot has too many entries.");
        }

        var entries = new List<ToonBlockEntry>(count);
        for (var i = 0; i < count; i++)
        {
            var isRemove = reader.Read(1) != 0;
            var region = (byte)reader.Read(8);
            var programId = (uint)reader.Read(32);
            var realm = (uint)reader.Read(32);
            var name = DecodeGeneratedUtf8(reader, lengthBits: 7, minimumBytes: 2, maximumBytes: 100, maximumCharacters: 25);
            entries.Add(new ToonBlockEntry(new ToonFullName(region, programId, realm, name), isRemove));
        }

        bool? complete = reader.Read(1) != 0 ? reader.Read(1) != 0 : null;
        return new ToonBlockNotifyRecord(entries, complete);
    }

    public static FriendsListRecord DecodeFriendsList(BitReader reader)
    {
        bool? complete = reader.Read(1) != 0 ? reader.Read(1) != 0 : null;
        var count = (int)reader.Read(7);
        if (count > 64)
        {
            throw new InvalidOperationException("Friends snapshot has too many updates.");
        }

        var updates = new List<FriendUpdate>(count);
        for (var i = 0; i < count; i++)
        {
            var operationIndex = reader.Read(2);
            if (operationIndex == 1)
            {
                FriendIdentity identity = reader.Read(1) == 0
                    ? new FriendIdentity.Account((uint)reader.Read(32))
                    : DecodeToonHandleIdentity(reader);
                updates.Add(new FriendUpdate(SocialOperation.Remove, new FriendEntry(identity, null, null, null, null, null)));
                continue;
            }

            var operation = operationIndex switch
            {
                0 => SocialOperation.Add,
                2 => SocialOperation.Modify,
                _ => throw new InvalidOperationException("Friends snapshot has an unknown update choice."),
            };
            updates.Add(new FriendUpdate(operation, DecodeFriendContainer(reader)));
        }

        return new FriendsListRecord(updates, complete);
    }

    public static ToonsOfFriendsRecord DecodeToonsOfFriends(BitReader reader)
    {
        var count = (int)reader.Read(7);
        if (count > 100)
        {
            throw new InvalidOperationException("Friend toon notification has too many entries.");
        }

        var entries = new List<FriendToon>(count);
        for (var i = 0; i < count; i++)
        {
            var region = (byte)reader.Read(8);
            var programId = (uint)reader.Read(32);
            var realm = (uint)reader.Read(32);
            var name = DecodeGeneratedUtf8(reader, lengthBits: 7, minimumBytes: 2, maximumBytes: 100, maximumCharacters: 25);
            var profileLabel = (uint)reader.Read(32);
            var profileId = reader.Read(64);
            var accountId = (uint)reader.Read(32);
            var profile = profileLabel != 0 || profileId != 0
                ? new PlayerTarget.ProfileRecordAddress(profileLabel, profileId)
                : null;
            entries.Add(new FriendToon(accountId, programId, profile, new ToonFullName(region, programId, realm, name)));
        }

        var complete = reader.Read(1) != 0;
        return new ToonsOfFriendsRecord(entries, complete);
    }

    private static FriendEntry DecodeFriendContainer(BitReader reader)
    {
        var choice = reader.Read(2);
        return choice switch
        {
            0 => DecodeFriendCharacter(reader),
            1 => DecodeFriendAccount(reader),
            2 => DecodeFriendPersistentPresenceUpdate(reader),
            _ => throw new InvalidOperationException("Friend update contains an unknown container choice."),
        };
    }

    private static FriendEntry DecodeFriendCharacter(BitReader reader)
    {
        var identity = DecodeToonHandleIdentity(reader);
        var displayName = DecodeGeneratedUtf8(reader, lengthBits: 7, minimumBytes: 2, maximumBytes: 100, maximumCharacters: 25);
        var profile = DecodeProfileRecordAddress(reader);
        var note = DecodeOptionalGeneratedUtf8(reader, lengthBits: 9, minimumBytes: 0, maximumBytes: 508, maximumCharacters: 127);
        return new FriendEntry(identity, displayName, null, note, profile, null);
    }

    private static FriendEntry DecodeFriendAccount(BitReader reader)
    {
        var accountId = (uint)reader.Read(32);
        var fullName = reader.Read(1) != 0 ? DecodeAccountFullName(reader) : null;
        var displayName = DecodeOptionalGeneratedUtf8(reader, lengthBits: 7, minimumBytes: 0, maximumBytes: 108, maximumCharacters: 27);
        var profile = DecodeProfileRecordAddress(reader);
        DiscardCustomMessage(reader);
        var note = DecodeOptionalGeneratedUtf8(reader, lengthBits: 9, minimumBytes: 0, maximumBytes: 508, maximumCharacters: 127);
        ReadS32(reader); // last_online — not carried on FriendEntry, matching upstream.
        reader.Read(64); // account_serial — discarded, matching upstream.
        reader.Read(32); // game_account_id — discarded, matching upstream.
        return new FriendEntry(new FriendIdentity.Account(accountId), displayName, fullName, note, profile, null);
    }

    /// <summary>AccountFullName: given name, then surname, each an 8-bit byte count and byte-aligned UTF-8. Joined with a space for display; either half may be empty.</summary>
    private static string DecodeAccountFullName(BitReader reader)
    {
        var given = DecodeGeneratedUtf8(reader, lengthBits: 8, minimumBytes: 0, maximumBytes: 255, maximumCharacters: 255);
        var surname = DecodeGeneratedUtf8(reader, lengthBits: 8, minimumBytes: 0, maximumBytes: 255, maximumCharacters: 255);
        return $"{given} {surname}".Trim();
    }

    private static FriendEntry DecodeFriendPersistentPresenceUpdate(BitReader reader)
    {
        var accountId = (uint)reader.Read(32);
        DiscardCustomMessage(reader);
        ReadS32(reader); // last_online — not carried on FriendEntry, matching upstream.
        return new FriendEntry(new FriendIdentity.Account(accountId), null, null, null, null, null);
    }

    /// <summary>Battlenet::Toon::Handle, reached from a friend identity. Same field order as <see cref="ToonRecordDecoder"/>'s already-verified convention (program, region, realm, id).</summary>
    private static FriendIdentity.Character DecodeToonHandleIdentity(BitReader reader)
    {
        var programId = (uint)reader.Read(32);
        var region = (byte)reader.Read(8);
        var realm = (uint)reader.Read(32);
        var id = reader.Read(64);
        return new FriendIdentity.Character(programId, region, realm, id);
    }

    /// <summary>Battlenet::Profile::RecordAddress — a plain (label, id) pair, no presence bit; matches <see cref="PlayerTarget.ProfileRecordAddress"/>'s already-established layout.</summary>
    private static PlayerTarget.ProfileRecordAddress DecodeProfileRecordAddress(BitReader reader)
    {
        var label = (uint)reader.Read(32);
        var id = reader.Read(64);
        return new PlayerTarget.ProfileRecordAddress(label, id);
    }

    /// <summary>Battlenet::Presence::CustomMessage — read and discarded; FriendEntry has no field for it, matching upstream's own FriendEntry shape.</summary>
    private static void DiscardCustomMessage(BitReader reader)
    {
        if (reader.Read(1) == 0)
        {
            return;
        }

        ReadS32(reader); // timestamp
        DecodeGeneratedUtf8(reader, lengthBits: 9, minimumBytes: 0, maximumBytes: 508, maximumCharacters: 127); // text
    }

    private static int ReadS32(BitReader reader) => unchecked((int)(reader.Read(32) ^ 0x8000_0000UL));

    private static string? DecodeOptionalGeneratedUtf8(BitReader reader, int lengthBits, int minimumBytes, int maximumBytes, int maximumCharacters) =>
        reader.Read(1) != 0 ? DecodeGeneratedUtf8(reader, lengthBits, minimumBytes, maximumBytes, maximumCharacters) : null;

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
