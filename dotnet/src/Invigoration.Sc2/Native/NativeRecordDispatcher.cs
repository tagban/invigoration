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
/// Covers every record the docs.bnet.cc Sunken record reference lists, plus
/// chat invites, incoming whispers and the Toon 18 reward catalog, which it
/// doesn't. A record whose (slot, command) isn't recognized throws rather
/// than silently skipping — per <see cref="RecordStream"/>'s remarks, an
/// unrecognized route can't be skipped safely without knowing its bit width,
/// so surfacing it loudly is the only safe option.
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

    /// <summary>Chat 30: Battle.net echoing a whisper we sent. Its peer is who we whispered.</summary>
    public sealed record WhisperEcho(ChatWhisperRecord Value) : NativeChatRecord;

    public sealed record Join(ChatJoinRecord Value) : NativeChatRecord;

    public sealed record FriendsList(FriendsListRecord Value) : NativeChatRecord;

    public sealed record ToonsOfFriends(ToonsOfFriendsRecord Value) : NativeChatRecord;

    public sealed record ToonBlocks(ToonBlockNotifyRecord Value) : NativeChatRecord;

    public sealed record ToonSelected(ToonSelectedRecord Value) : NativeChatRecord;

    /// <summary>Club slot 13, command 46: our character's clubs (GetToonClubsResponse).</summary>
    public sealed record ToonClubs(ToonClubsRecord Value) : NativeChatRecord;

    /// <summary>Club slot 13, command 54: a club invitation, or an answer to one (InviteAction).</summary>
    public sealed record ClubInvite(ClubInviteRecord Value) : NativeChatRecord;

    /// <summary>Club slot 13, command 50: members joined, left, or changed rank or status.</summary>
    public sealed record ClubMemberChanges(ClubMemberChangesRecord Value) : NativeChatRecord;

    /// <summary>Club slot 13, command 57, at sign-in: the rules for club names and tags.</summary>
    public sealed record ClubSettings(ClubSettingsRecord Value) : NativeChatRecord;

    /// <summary>Club slot 13, command 45: a page of a club's members.</summary>
    public sealed record ClubRoster(ClubRosterRecord Value) : NativeChatRecord;

    /// <summary>Club slot 13, command 49: changes to clubs we follow (a full sync right after subscribing).</summary>
    public sealed record ClubChanges(ClubChangesRecord Value) : NativeChatRecord;

    /// <summary>Profile slot 14, command 2: names for character handles.</summary>
    public sealed record ToonNames(ToonNamesRecord Value) : NativeChatRecord;

    public sealed record ToonList(ToonListRecord Value) : NativeChatRecord;

    /// <summary>
    /// Connection/GameSiteInfo (a regional game-server catalog, e.g. "US10-S2",
    /// "ORD1-S2", "AU1-S2", "SA1-S2", "US3", "SG1" in two live captures), sent
    /// unprompted right after Resume/EnableEncryption. Read field by field by
    /// <see cref="ConnectionRecordDecoder.SkipGameSiteInfo"/>; the contents aren't kept.
    /// </summary>
    public sealed record Sc2ServerCatalog : NativeChatRecord;

    /// <summary>
    /// Toon/Welcome (Toon slot, command 10), sent once a character is selected.
    /// Read field by field by <see cref="StartupRecordDecoder.SkipToonWelcome"/>;
    /// only the number of achievement/unlock entries it listed is kept.
    /// </summary>
    public sealed record Sc2ToonWelcomeUnlocks(int UnlockCount) : NativeChatRecord;

    /// <summary>
    /// Toon slot, command 18 — sent right after Welcome. Readable strings in a
    /// live capture ("WarChestSeason1TerranTier1Bundle", "...ZergTier1Bundle",
    /// "...ProtossTier1Bundle", repeating per season/tier) identify this as a
    /// seasonal reward/"War Chest" bundle catalog, not anything chat-relevant.
    /// Not reverse-engineered, and not in the record reference either, so it's
    /// still consumed wholesale: the one guessed skip left in this dispatcher.
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

    /// <summary>Connection 10: Battle.net's keepalive. Answer it at once with <see cref="ConnectionCommands.Pong"/>, passing <see cref="Timestamp"/> back unchanged.</summary>
    public sealed record Ping(byte[]? Timestamp) : NativeChatRecord;

    /// <summary>Connection 12: Battle.net's answer to our own <see cref="ConnectionCommands.Ping"/>.</summary>
    public sealed record Pong(byte[]? Timestamp) : NativeChatRecord;

    /// <summary>Chat 22: the public channel list Battle.net sends back for Chat 21. Numbers only, no names.</summary>
    public sealed record PublicChannelList(IReadOnlyList<(byte A, ushort B, uint C)> Entries) : NativeChatRecord;

    /// <summary>Presence 1 (FieldSpecAnnounce): presence field definitions. Pass to <see cref="PresenceTracker.Announce"/>.</summary>
    public sealed record PresenceFields(PresenceFieldsRecord Value) : NativeChatRecord;

    /// <summary>Presence 0 (PresenceUpdateNotify): new field values for one presence. Pass to <see cref="PresenceTracker.Apply"/>.</summary>
    public sealed record PresenceUpdate(PresenceUpdateRecord Value) : NativeChatRecord;

    /// <summary>Profile 0: one answer to a <see cref="ChatCommands.ProfileReadRequest(uint, PlayerTarget.ProfileRecordAddress)"/>. Pass to <see cref="PortraitResolver.Complete"/>.</summary>
    public sealed record ProfileRead(ProfileReadRecord Value) : NativeChatRecord;

    /// <summary>Presence 10: the result of a temporary presence request.</summary>
    public sealed record TemporaryPresenceResult(ushort Result) : NativeChatRecord;

    /// <summary>
    /// A record read to its exact end only to stay in step with the stream:
    /// presence, billing, season, profile, message frames and the like. Its
    /// route is kept for logging; its contents aren't.
    /// </summary>
    public sealed record Sc2Consumed(byte Slot, byte Command) : NativeChatRecord;

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
    public static NativeChatRecord Decode(byte commandId, byte? serviceSlot, BitReader reader)
    {
        var recordStart = reader.Position - (serviceSlot is null ? 7 : 11);
        return (serviceSlot, commandId) switch
        {
            (null, _) => new NativeChatRecord.Sc2CommandAck((ushort)reader.Read(9)),

            (ConnectionSlot, ConnectionBoomCommand) => throw new NativeServerRejectedException((ushort)reader.Read(16)),
            (ConnectionSlot, 3) => Consumed(reader, ConnectionRecordDecoder.SkipCommand3, serviceSlot, commandId),
            (ConnectionSlot, ConnectionCommands.PingCommand) => new NativeChatRecord.Ping(ConnectionRecordDecoder.DecodeKeepalive(reader)),
            (ConnectionSlot, 11) => Consumed(reader, ConnectionRecordDecoder.SkipCommand11, serviceSlot, commandId),
            (ConnectionSlot, ConnectionCommands.PongCommand) => new NativeChatRecord.Pong(ConnectionRecordDecoder.DecodeKeepalive(reader)),
            (ConnectionSlot, 13) => Consumed(reader, ConnectionRecordDecoder.SkipMessageFrame, serviceSlot, commandId),
            (ConnectionSlot, GameSiteInfoCommand) => SkipGameSiteInfo(reader),

            (ChatCommands.ChatSlot, 1) => new NativeChatRecord.Membership(MembershipChangeDecoder.Decode(reader)),
            (ChatCommands.ChatSlot, 4) => new NativeChatRecord.Invite(ChatRecordDecoder.DecodeChatInvite(reader)),
            (ChatCommands.ChatSlot, 11) => new NativeChatRecord.Message(ChatRecordDecoder.DecodeChatMessage(reader)),
            (ChatCommands.ChatSlot, 19) => new NativeChatRecord.Whisper(ChatRecordDecoder.DecodeChatWhisper(reader)),
            (ChatCommands.ChatSlot, 22) => new NativeChatRecord.PublicChannelList(StartupRecordDecoder.DecodeChannelListResponse(reader)),
            (ChatCommands.ChatSlot, 24) => Consumed(reader, StartupRecordDecoder.SkipChannelCategories, serviceSlot, commandId),
            (ChatCommands.ChatSlot, 26) => Consumed(reader, StartupRecordDecoder.SkipChannelMemberCounts, serviceSlot, commandId),
            (ChatCommands.ChatSlot, 27) => new NativeChatRecord.Join(ChatRecordDecoder.DecodeChatJoin(reader)),
            (ChatCommands.ChatSlot, 30) => new NativeChatRecord.WhisperEcho(ChatRecordDecoder.DecodeChatWhisper(reader)),

            (FriendsSlot, FriendsToonsCommand) => new NativeChatRecord.ToonsOfFriends(FriendsRecordDecoder.DecodeToonsOfFriends(reader)),
            (FriendsSlot, FriendsListCommand) => new NativeChatRecord.FriendsList(FriendsRecordDecoder.DecodeFriendsList(reader)),
            (FriendsSlot, 31) => Consumed(reader, StartupRecordDecoder.SkipAccountBlocks, serviceSlot, commandId),
            (FriendsSlot, FriendsToonBlockCommand) => new NativeChatRecord.ToonBlocks(FriendsRecordDecoder.DecodeToonBlockNotify(reader)),

            (PresenceSlot, 0) => new NativeChatRecord.PresenceUpdate(PresenceRecordDecoder.DecodePresenceUpdate(reader)),
            (PresenceSlot, 1) => new NativeChatRecord.PresenceFields(PresenceRecordDecoder.DecodeFieldSpecAnnounce(reader)),
            (PresenceSlot, 2) => Consumed(reader, r => r.Read(1), serviceSlot, commandId),
            (PresenceSlot, 3) => Consumed(reader, StartupRecordDecoder.SkipPresenceStatistics, serviceSlot, commandId),
            (PresenceSlot, 4) => Consumed(reader, StartupRecordDecoder.SkipTemporaryPresenceRequest, serviceSlot, commandId),
            (PresenceSlot, 10) => new NativeChatRecord.TemporaryPresenceResult((ushort)reader.Read(16)),

            (ChatCommands.ToonSlot, 0) => new NativeChatRecord.ToonList(ToonRecordDecoder.DecodeToonList(reader)),
            (ChatCommands.ToonSlot, 6) => new NativeChatRecord.ToonSelected(ToonRecordDecoder.DecodeToonSelected(reader)),
            (ChatCommands.ToonSlot, ToonWelcomeCommand) => new NativeChatRecord.Sc2ToonWelcomeUnlocks(StartupRecordDecoder.SkipToonWelcome(reader)),
            (ChatCommands.ToonSlot, 13) => Consumed(reader, StartupRecordDecoder.SkipBillingUpdate, serviceSlot, commandId),
            (ChatCommands.ToonSlot, 14) => Consumed(reader, _ => { }, serviceSlot, commandId),
            (ChatCommands.ToonSlot, ToonRewardCatalogCommand) => ConsumeRestOfRecord(reader, new NativeChatRecord.Sc2RewardCatalog()),

            (ChatCommands.CacheSlot, CacheGetStreamItemsCommand) => SkipCacheStreamItems(reader),
            (S2MasterSlot, 27) => Consumed(reader, StartupRecordDecoder.SkipCurrentSeason, serviceSlot, commandId),
            (PartySlot, 0) => Consumed(reader, r => r.SkipToRecordByte(recordStart, 18), serviceSlot, commandId),
            (S2MapsSlot, ClubCommands.GetToonClubsCommand) => new NativeChatRecord.ToonClubs(Club(() => ClubCommands.DecodeToonClubs(reader), commandId)),
            (S2MapsSlot, ClubCommands.InviteActionCommand) => new NativeChatRecord.ClubInvite(Club(() => ClubCommands.DecodeInviteAction(reader), commandId)),
            (S2MapsSlot, ClubCommands.GetRosterCommand) => new NativeChatRecord.ClubRoster(Club(() => ClubCommands.DecodeRoster(reader), commandId)),
            (ProfileSlot, ClubCommands.ResolveToonNamesCommand) => new NativeChatRecord.ToonNames(Club(() => ClubCommands.DecodeToonNames(reader), commandId, ProfileSlot)),
            (S2MapsSlot, ClubCommands.ClubChangeNotificationCommand) => new NativeChatRecord.ClubChanges(Club(() => ClubCommands.DecodeClubChanges(reader), commandId)),
            (S2MapsSlot, ClubCommands.MemberChangeNotificationCommand) => new NativeChatRecord.ClubMemberChanges(Club(() => ClubCommands.DecodeMemberChanges(reader), commandId)),
            (S2MapsSlot, ClubCommands.ClubSettingsCommand) => new NativeChatRecord.ClubSettings(Club(() => ClubCommands.DecodeClubSettings(reader), commandId)),
            (ProfileSlot, ChatCommands.ProfileReadCommand) => new NativeChatRecord.ProfileRead(ProfileRecordDecoder.DecodeProfileRead(reader)),
            (ProfileSlot, 4) => Consumed(reader, StartupRecordDecoder.SkipProfileSettings, serviceSlot, commandId),

            _ => throw new UnknownNativeRecordException(serviceSlot, commandId),
        };
    }

    /// <summary>
    /// A club record read by schema. If its layout turns out wrong it's treated as unknown (the
    /// buffer is dropped) rather than ending the session: clubs are optional.
    /// </summary>
    private static T Club<T>(Func<T> decode, byte commandId, byte slot = S2MapsSlot)
    {
        try
        {
            return decode();
        }
        catch (Exception ex) when (ex is Bsn.BsnException or InvalidCastException or NullReferenceException)
        {
            throw new UnknownNativeRecordException(slot, commandId);
        }
    }

    private static NativeChatRecord.Sc2Consumed Consumed(BitReader reader, Action<BitReader> read, byte? slot, byte command)
    {
        read(reader);
        return new NativeChatRecord.Sc2Consumed(slot!.Value, command);
    }

    private const byte PresenceSlot = 4;
    private const byte S2MasterSlot = 10;
    private const byte PartySlot = 12;
    private const byte S2MapsSlot = 13;
    private const byte ProfileSlot = ChatCommands.ProfileSlot;

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

    private static NativeChatRecord.Sc2ServerCatalog SkipGameSiteInfo(BitReader reader)
    {
        ConnectionRecordDecoder.SkipGameSiteInfo(reader);
        return new NativeChatRecord.Sc2ServerCatalog();
    }

    /// <summary>core/src/native/protocol.rs: TOON_WELCOME_COMMAND.</summary>
    private const byte ToonWelcomeCommand = 10;

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

/// <summary>A record whose (slot, command) no decoder knows. Its length can't be known either.</summary>
public sealed class UnknownNativeRecordException(byte? slot, byte command)
    : InvalidOperationException($"No decoder registered for native record slot={slot} command={command}.")
{
    public byte? Slot { get; } = slot;

    public byte Command { get; } = command;
}
