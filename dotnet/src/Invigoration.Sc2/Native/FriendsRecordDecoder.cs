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
/// ChatBootstrap alongside FriendsList/ToonsOfFriends. Field widths (7-bit
/// array length capped at 64, and each entry's Battlenet::Toon::FullName +
/// 1-bit Add/Remove choice) were confirmed bit-exact against a real captured
/// record using the extracted retail schema (types 2724/2725/2729/2732/1053)
/// — see the "extract-bsn-metadata" tool and decode_hex example in
/// ncarrillo/superiority's repo. Notably, Battlenet::Toon::Name's length
/// field here is only 5 bits (bias +2, cap 25 chars) — narrower than the
/// 7-bit display-name field used elsewhere in this file — and its bytes
/// must be read byte-aligned (<see cref="DecodeGeneratedUtf8"/> already does
/// this), not as a raw unaligned bit run.
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
/// One field could not be recovered this way: FriendContainer5::Account's
/// optional m_fullName (an AccountFullName — given/surname) is read upstream
/// via the fully generic, schema-blob-driven codec path
/// (Protocol::codec().decode_reflected_traced_from), not a hand-traced
/// sequence of fixed-width reads — unlike everything else in this file,
/// there is no fixed bit layout to port for it. <see cref="DecodeFriendAccount"/>
/// throws if that field's presence bit is ever set, rather than silently
/// desyncing the stream past an unknown-width structure.
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
            var region = (byte)reader.Read(8);
            var programId = (uint)reader.Read(32);
            var realm = (uint)reader.Read(32);
            var name = DecodeGeneratedUtf8(reader, lengthBits: 5, minimumBytes: 2, maximumBytes: 33, maximumCharacters: 25);
            var isRemove = reader.Read(1) != 0;
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
        if (reader.Read(1) != 0)
        {
            throw new NotSupportedException(
                "This friend has a full name (AccountFullName) set, and that field's wire layout " +
                "is only known via SC2's embedded runtime schema — see this decoder's remarks.");
        }

        var displayName = DecodeOptionalGeneratedUtf8(reader, lengthBits: 7, minimumBytes: 0, maximumBytes: 108, maximumCharacters: 27);
        var profile = DecodeProfileRecordAddress(reader);
        DiscardCustomMessage(reader);
        var note = DecodeOptionalGeneratedUtf8(reader, lengthBits: 9, minimumBytes: 0, maximumBytes: 508, maximumCharacters: 127);
        ReadS32(reader); // last_online — not carried on FriendEntry, matching upstream.
        reader.Read(64); // account_serial — discarded, matching upstream.
        reader.Read(32); // game_account_id — discarded, matching upstream.
        return new FriendEntry(new FriendIdentity.Account(accountId), displayName, null, note, profile, null);
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
