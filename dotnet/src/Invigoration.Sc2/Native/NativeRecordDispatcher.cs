using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

/// <summary>
/// One decoded native ("Sunken") record, routed to whichever typed decoder
/// its (service slot, command) pair identifies. This is the connective layer
/// between <see cref="RecordStream.TryDecodeRecord{T}"/> (which needs a
/// caller-supplied decode function per record) and the individual hand-rolled
/// decoders in this namespace — it does not itself add any new wire-format
/// knowledge.
///
/// Deliberately narrow: only the record types this project can already
/// decode (see <see cref="ChatRecordDecoder"/>, <see cref="MembershipChangeDecoder"/>,
/// <see cref="FriendsRecordDecoder"/>, <see cref="ToonRecordDecoder"/>) are
/// routed here. A record whose (slot, command) isn't recognized throws
/// rather than silently skipping — per <see cref="RecordStream"/>'s remarks,
/// an unrecognized route can't be skipped safely without knowing its bit
/// width, so surfacing it loudly is the only safe option.
///
/// What this layer does NOT yet do (left for whoever picks this up next):
/// toon select / channel join sequencing, roster/presence tracking, or
/// resolving a member handle to a display name — those need their own
/// stateful session type (mirroring core/src/chat/session.rs's LiveChat),
/// not just record decoding.
/// </summary>
public abstract record NativeChatRecord
{
    private NativeChatRecord()
    {
    }

    public sealed record Membership(MembershipChangeNotifyRecord Value) : NativeChatRecord;

    public sealed record Invite(ChatInviteRecord Value) : NativeChatRecord;

    public sealed record Message(ChatMessageRecord Value) : NativeChatRecord;

    public sealed record Whisper(ChatWhisperRecord Value) : NativeChatRecord;

    public sealed record Join(ChatJoinRecord Value) : NativeChatRecord;

    public sealed record FriendsList(FriendsListRecord Value) : NativeChatRecord;

    public sealed record ToonsOfFriends(ToonsOfFriendsRecord Value) : NativeChatRecord;

    public sealed record ToonBlocks(ToonBlockNotifyRecord Value) : NativeChatRecord;

    public sealed record ToonSelected(ToonSelectedRecord Value) : NativeChatRecord;

    public sealed record ToonList(ToonListRecord Value) : NativeChatRecord;

    /// <summary>
    /// Connection/GameSiteInfo (a regional game-server catalog, e.g. "US10-S2",
    /// "ORD1-S2", "AU1-S2", "SA1-S2", "US3", "SG1" in two live captures) — sent
    /// unprompted right after Resume/EnableEncryption. Upstream's own reference
    /// only decodes this via the actual SC2 client's embedded schema blob, which
    /// isn't available here, so its exact field layout (beyond the site names
    /// themselves) isn't modeled — see <see cref="NativeRecordDispatcher.Decode"/>'s
    /// remarks on this record's empirical, not-fully-understood skip.
    /// </summary>
    public sealed record Sc2ServerCatalog : NativeChatRecord;

    /// <summary>
    /// Toon/Welcome (Toon slot, command 10) — sent once a character is
    /// selected. Contains, among other things, a per-account achievement/
    /// unlock array: each entry is a fixed 41 bytes starting with an "unlk"
    /// (0x756E6C6B) marker, self-identifying enough to scan without knowing
    /// the array's declared count. See <see cref="NativeRecordDispatcher.Decode"/>'s
    /// remarks — the array itself decodes cleanly, but what follows it in
    /// the record isn't understood yet, so this always ends by throwing once
    /// the array ends, surfacing the rest of the record for further capture.
    /// </summary>
    public sealed record Sc2ToonWelcomeUnlocks(int UnlockCount) : NativeChatRecord;

    /// <summary>
    /// Toon slot, command 18 — sent right after Welcome. Readable strings in a
    /// live capture ("WarChestSeason1TerranTier1Bundle", "...ZergTier1Bundle",
    /// "...ProtossTier1Bundle", repeating per season/tier) identify this as a
    /// seasonal reward/"War Chest" bundle catalog, not anything chat-relevant.
    /// Not reverse-engineered — consumed wholesale like <see cref="Sc2ToonWelcomeUnlocks"/>'s tail.
    /// </summary>
    public sealed record Sc2RewardCatalog : NativeChatRecord;

    /// <summary>
    /// Cache slot, command 9 — the response to a CacheGetStreamItems request
    /// (see <see cref="ChatCommands.CacheGetStreamItems"/>). Each item only
    /// carries a 40-byte content handle that points at a Blizzard CDN blob
    /// (the real catalog XML lives there, resolved out-of-band); fetching
    /// and parsing that isn't needed for chat, so this decoder exists purely
    /// to consume the record's exact bit width and keep the stream framed
    /// correctly, not to expose any catalog data.
    /// </summary>
    public sealed record Sc2CacheCatalogResponse : NativeChatRecord;

    /// <summary>
    /// A placeholder for a record type nobody's added a decoder for, consumed
    /// wholesale by the caller (see <c>BotEngine.Sc2.cs</c>'s lenient pre-join
    /// decoding) rather than by this dispatcher — unlike
    /// <see cref="Sc2ServerCatalog"/>/<see cref="Sc2RewardCatalog"/>, which are
    /// specific known routes this dispatcher itself knows how to skip.
    /// </summary>
    public sealed record Sc2UnknownStartupRecord(byte? Slot, byte Command) : NativeChatRecord;

    /// <summary>
    /// A "command response" record — the 7-bit-header variant (no service
    /// slot) used to acknowledge a request the client sent, e.g. the two
    /// CacheGetStreamItems bootstrap requests <c>BotEngine.Sc2.cs</c> sends
    /// right after connecting. Just a 9-bit result code; core/src/native/stream.rs
    /// decodes this via its own dedicated decode_command_response path rather
    /// than routing it through the normal per-(slot,command) table the way
    /// every other record here is. This project didn't model that distinction
    /// at all until now — these were silently falling into the pre-join
    /// lenient skip (which discards the *entire* remaining buffer, not just
    /// this one small record), which is unsafe if anything else happened to
    /// be buffered right behind it.
    /// </summary>
    public sealed record Sc2CommandAck(ushort Result) : NativeChatRecord;
}

public static class NativeRecordDispatcher
{
    /// <summary>
    /// Decodes exactly one record given its already-parsed routing header.
    /// Pass this to <see cref="RecordStream.TryDecodeRecord{T}"/> (it matches
    /// that method's decode-function shape) rather than calling it directly,
    /// so buffering/underrun handling stays centralized there.
    /// </summary>
    public static NativeChatRecord Decode(byte commandId, byte? serviceSlot, BitReader reader) =>
        (serviceSlot, commandId) switch
        {
            (null, _) => new NativeChatRecord.Sc2CommandAck((ushort)reader.Read(9)),
            (ChatCommands.ChatSlot, 1) => new NativeChatRecord.Membership(MembershipChangeDecoder.Decode(reader)),
            (ChatCommands.ChatSlot, 4) => new NativeChatRecord.Invite(ChatRecordDecoder.DecodeChatInvite(reader)),
            (ChatCommands.ChatSlot, 11) => new NativeChatRecord.Message(ChatRecordDecoder.DecodeChatMessage(reader)),
            (ChatCommands.ChatSlot, 19) => new NativeChatRecord.Whisper(ChatRecordDecoder.DecodeChatWhisper(reader)),
            (ChatCommands.ChatSlot, 27) => new NativeChatRecord.Join(ChatRecordDecoder.DecodeChatJoin(reader)),
            (ChatCommands.ChatSlot, 30) => new NativeChatRecord.Whisper(ChatRecordDecoder.DecodeChatWhisper(reader)),
            (FriendsSlot, FriendsListCommand) => new NativeChatRecord.FriendsList(FriendsRecordDecoder.DecodeFriendsList(reader)),
            (FriendsSlot, FriendsToonsCommand) => new NativeChatRecord.ToonsOfFriends(FriendsRecordDecoder.DecodeToonsOfFriends(reader)),
            (FriendsSlot, FriendsToonBlockCommand) => new NativeChatRecord.ToonBlocks(FriendsRecordDecoder.DecodeToonBlockNotify(reader)),
            (ChatCommands.ToonSlot, 6) => new NativeChatRecord.ToonSelected(ToonRecordDecoder.DecodeToonSelected(reader)),
            (ChatCommands.ToonSlot, 0) => new NativeChatRecord.ToonList(ToonRecordDecoder.DecodeToonList(reader)),
            (ConnectionSlot, ConnectionBoomCommand) => throw new NativeServerRejectedException((ushort)reader.Read(16)),
            (ConnectionSlot, GameSiteInfoCommand) => SkipGameSiteInfo(reader),
            (ChatCommands.ToonSlot, ToonWelcomeCommand) => SkipToonWelcome(reader),
            (ChatCommands.ToonSlot, ToonRewardCatalogCommand) => ConsumeRestOfRecord(reader, new NativeChatRecord.Sc2RewardCatalog()),
            (ChatCommands.CacheSlot, CacheGetStreamItemsCommand) => SkipCacheStreamItems(reader),
            _ => throw new InvalidOperationException(
                $"No decoder registered for native record slot={serviceSlot} command={commandId}."),
        };

    /// <summary>Battlenet::Friends' RPC service slot. core/src/native/protocol.rs: FRIENDS_SLOT.</summary>
    private const byte FriendsSlot = 3;

    /// <summary>core/src/native/protocol.rs: FRIENDS_LIST_COMMAND — routes to FriendsListNotify5.</summary>
    private const byte FriendsListCommand = 30;

    /// <summary>core/src/native/protocol.rs: FRIENDS_TOONS_COMMAND — routes to ToonsOfFriendsNotify.</summary>
    private const byte FriendsToonsCommand = 6;

    /// <summary>Friends slot, command 33 — Battlenet::Client::Friends::ToonBlockNotify, confirmed via the extracted retail schema (type #2724).</summary>
    private const byte FriendsToonBlockCommand = 33;

    /// <summary>core/src/native/protocol.rs: CONNECTION_SLOT.</summary>
    private const byte ConnectionSlot = 1;

    /// <summary>core/src/native/protocol.rs: CONNECTION_BOOM_COMMAND — the server's explicit "here's why I'm disconnecting you" message. Matches the same decode already used during the Resume handshake in SunkenClient.cs, just also wired in here for the ongoing post-handshake receive loop.</summary>
    private const byte ConnectionBoomCommand = 1;

    /// <summary>core/src/native/protocol.rs: CONNECTION_GAME_SITE_INFO_COMMAND.</summary>
    private const byte GameSiteInfoCommand = 14;

    /// <summary>
    /// Reads (and discards) a GameSiteInfo record — Battlenet::Client::Connection::GameSiteInfo,
    /// a regional game-server catalog (m_externalIp4Addr: {address, port}, then
    /// m_siteData: an array of {name, optional address/port} site entries — site
    /// codes like "US10-S2"/"ORD1-S2"/"AU1-S2"/"SA1-S2"/"US3"/"SG1" were directly
    /// legible in a live capture). Confirmed bit-exact (logical_bits=712, i.e. 701
    /// bits of payload after this method's caller already consumed the 11-bit
    /// routing header) against a real live capture using the actual game client's
    /// own embedded BSN schema — see the "extract-bsn-metadata" tool and
    /// decode_hex example in ncarrillo/superiority's repo, which can decode any
    /// captured record exactly given a metadata blob pulled from a real SC2.exe.
    /// This project still doesn't carry that schema or a generic codec, so this
    /// reads the confirmed bit count and discards it rather than modeling every
    /// field — content isn't needed for chat.
    /// </summary>
    private static NativeChatRecord.Sc2ServerCatalog SkipGameSiteInfo(BitReader reader)
    {
        const int totalRecordBits = 712;
        const int routingHeaderBits = 11;
        reader.Read(totalRecordBits - routingHeaderBits);
        return new NativeChatRecord.Sc2ServerCatalog();
    }

    /// <summary>core/src/native/protocol.rs: TOON_WELCOME_COMMAND.</summary>
    private const byte ToonWelcomeCommand = 10;

    private static readonly byte[] ToonUnlockMarker = [0x75, 0x6E, 0x6C, 0x6B]; // "unlk"

    /// <summary>
    /// Reads Toon/Welcome's fixed 109-byte pre-array header (byte-aligned from
    /// right after the routing header — empirically measured from a live
    /// capture, not derived), then consumes 41-byte "unlk"-marked achievement
    /// entries for as long as they keep appearing (confirmed against a real
    /// 123-entry account). What follows the array — deep, unrelated-looking
    /// structured data (map/mod/mastery metadata; a live capture showed Java
    /// class-file magic numbers and what looks like a regex pattern) — is not
    /// reverse-engineered at all: it's consumed wholesale, on the bet that
    /// Battle.net delivers this whole record as one logical chunk. If that
    /// bet is ever wrong, the *next* record will fail to decode — loudly,
    /// same as every other route in this dispatcher — rather than silently
    /// desyncing further.
    /// </summary>
    private static NativeChatRecord.Sc2ToonWelcomeUnlocks SkipToonWelcome(BitReader reader)
    {
        const int headerBytes = 109;
        const int unlockEntryBytesAfterMarker = 37;
        reader.Align();
        reader.ReadBytes(headerBytes, aligned: true);

        var count = 0;
        while (reader.RemainingBits >= 32)
        {
            var marker = reader.ReadBytes(4, aligned: true);
            if (!marker.AsSpan().SequenceEqual(ToonUnlockMarker) ||
                reader.RemainingBits < unlockEntryBytesAfterMarker * 8)
            {
                break;
            }

            reader.ReadBytes(unlockEntryBytesAfterMarker, aligned: true);
            count++;
        }

        return ConsumeRestOfRecord(reader, new NativeChatRecord.Sc2ToonWelcomeUnlocks(count));
    }

    /// <summary>core/src/native/protocol.rs doesn't name this one (not in the version of the reference this project has) — Toon slot, command 18, empirically a War Chest/reward-bundle catalog.</summary>
    private const byte ToonRewardCatalogCommand = 18;

    /// <summary>
    /// Discards whatever's left in the currently-buffered record without
    /// attempting to understand it — used only for records confirmed to be
    /// large, connect-time-only "catalog" blobs (achievements, reward
    /// bundles) where guessing wrong just means the *next* record fails to
    /// decode loudly, not a silent runtime desync. Never use this for a
    /// route that could plausibly be a small, frequent, or chat-critical
    /// message — see <see cref="RecordStream"/>'s remarks on why an
    /// unrecognized route normally can't be skipped at all.
    /// </summary>
    private static T ConsumeRestOfRecord<T>(BitReader reader, T record)
    {
        if (reader.RemainingBits > 0)
        {
            reader.Read(reader.RemainingBits);
        }

        return record;
    }

    /// <summary>core/src/native/protocol.rs: CACHE_GET_STREAM_ITEMS_COMMAND.</summary>
    private const byte CacheGetStreamItemsCommand = 9;

    /// <summary>
    /// Reads (and discards) a CacheStreamItems response — a 6-bit item count
    /// (capped at 49, matching upstream's own sanity check; Battle.net paginates
    /// larger catalogs across multiple responses rather than exceeding this),
    /// each item being a 23-bit obfuscation selector + a byte-aligned 40-byte
    /// content handle + a sign-flipped int32 publication time, followed by a
    /// 32-bit token, 16-bit total-item-count, and 16-bit offset. Ported from
    /// core/src/native/decode.rs's cache_stream_items_with_provenance. See
    /// <see cref="NativeChatRecord.Sc2CacheCatalogResponse"/> for why the
    /// content handles themselves aren't resolved.
    /// </summary>
    private static NativeChatRecord.Sc2CacheCatalogResponse SkipCacheStreamItems(BitReader reader)
    {
        var count = (int)reader.Read(6);
        if (count > 49)
        {
            throw new InvalidOperationException("Cache stream response contains too many items.");
        }

        for (var i = 0; i < count; i++)
        {
            reader.Read(23); // wire_layout_selector (obfuscation), discarded.
            reader.ReadBytes(40, aligned: true); // content_handle, discarded.
            reader.Read(32); // publication_time (sign-flip int32), discarded.
        }

        reader.Read(32); // token
        reader.Read(16); // total_items
        reader.Read(16); // offset
        return new NativeChatRecord.Sc2CacheCatalogResponse();
    }
}
