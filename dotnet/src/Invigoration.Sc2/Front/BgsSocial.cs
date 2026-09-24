using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Front;

/// <summary>Value-equality key for a bgs.protocol.EntityId (the class itself compares by reference).</summary>
public readonly record struct BgsEntityKey(ulong High, ulong Low)
{
    public static BgsEntityKey From(EntityId id) => new(id.High, id.Low);

    public EntityId ToEntityId() => new() { High = High, Low = Low };

    public override string ToString() => $"{High:X16}/{Low}";
}

/// <summary>Value-equality form of <see cref="BgsPresenceFieldKey"/>; a missing unique_id counts as 0.</summary>
public readonly record struct BgsFieldKey(uint Program, uint Group, uint Field, ulong UniqueId = 0)
{
    public static BgsFieldKey From(BgsPresenceFieldKey key) => new(key.Program, key.Group, key.Field, key.UniqueId ?? 0);
}

/// <summary>
/// Battle.net-level presence field numbers, program "BN" (0x424E). Group 1 lives on the ACCOUNT
/// entity, group 2 on each GAME ACCOUNT entity.
///
/// Sourced from two open BGS server emulators that had to answer the real Diablo III client's
/// presence queries: mooege (DarkLotus/mooege, Core/MooNet/Accounts/Account.cs and GameAccount.cs,
/// 2012) and blizzless-diiis (BGS-Server/AccountsSystem/Account.cs and GameAccount.cs, D3 2.7.x).
/// Both agree on every number below except where noted. Neither is a capture of today's
/// Battle.net app, so the layout needs a live check; unknown fields stay reachable through
/// <see cref="BgsPresenceTracker.GetFields"/>.
/// </summary>
public static class BgsPresenceFields
{
    public const uint BattleNetProgram = 0x424E; // "BN"
    public const uint AccountGroup = 1;
    public const uint GameAccountGroup = 2;

    /// <summary>Real ID / full name (string). mooege "RealIDTagField".</summary>
    public const uint AccountFullName = 1;

    /// <summary>A bool "account online" in diiis's AccountOnlineField (commented out in mooege), but diiis also sends its string broadcast message on this number; the tracker reads it by value type.</summary>
    public const uint AccountOnlineOrBroadcast = 2;

    /// <summary>Game account ids, one entry per unique_id (EntityId values).</summary>
    public const uint AccountGameAccounts = 3;

    /// <summary>BattleTag, "Name#1234" (string).</summary>
    public const uint AccountBattleTag = 4;

    /// <summary>Last online time (int).</summary>
    public const uint AccountLastOnline = 6;

    /// <summary>Game account signed in (bool).</summary>
    public const uint GameAccountIsOnline = 1;

    /// <summary>Away status flags (int): 0x02 away, 0x04 busy (diiis AwayStatusFlag).</summary>
    public const uint GameAccountAwayStatus = 2;

    /// <summary>Program the game account is in (FourCC string, e.g. "S2", "S1", "BSAp", "Fen").</summary>
    public const uint GameAccountProgram = 3;

    /// <summary>Last online time (int).</summary>
    public const uint GameAccountLastOnline = 4;

    /// <summary>BattleTag (string).</summary>
    public const uint GameAccountBattleTag = 5;

    /// <summary>Game account name (string), e.g. "12345#1".</summary>
    public const uint GameAccountName = 6;

    /// <summary>Owning account's EntityId.</summary>
    public const uint GameAccountOwner = 7;

    /// <summary>Rich presence: a RichPresenceLocalizationKey {program, stream, localization_id} in message_value, i.e. an index into the game's own string tables, not text.</summary>
    public const uint GameAccountRichPresence = 8;

    /// <summary>AFK (bool).</summary>
    public const uint GameAccountAfk = 10;

    public const uint AwayFlag = 0x02;
    public const uint BusyFlag = 0x04;
}

/// <summary>Rich presence as Battle.net carries it: which string of which program's localization stream, not the text itself.</summary>
public sealed record BgsRichPresence(uint Program, uint Stream, uint LocalizationId)
{
    public string ProgramName => FourCc.Decode(Program);

    /// <summary>RichPresenceLocalizationKey {program = 1 (fixed32), stream = 2 (fixed32), localization_id = 3}.</summary>
    public static BgsRichPresence? Decode(byte[]? data)
    {
        if (data is null || data.Length == 0)
        {
            return null;
        }

        try
        {
            uint program = 0, stream = 0, id = 0;
            var r = new Protobuf.ProtoReader(data);
            while (r.HasMore)
            {
                var (field, type) = r.ReadTag();
                switch (field)
                {
                    case 1 when type == Protobuf.WireType.Fixed32: program = r.ReadFixed32(); break;
                    case 2 when type == Protobuf.WireType.Fixed32: stream = r.ReadFixed32(); break;
                    case 3 when type == Protobuf.WireType.Varint: id = (uint)r.ReadVarint(); break;
                    default: r.Skip(type); break;
                }
            }

            return new BgsRichPresence(program, stream, id);
        }
        catch (Exception ex) when (ex is IndexOutOfRangeException or ArgumentOutOfRangeException or InvalidOperationException)
        {
            return null;
        }
    }

    public byte[] Encode()
    {
        var w = new Protobuf.ProtoWriter();
        w.WriteFixed32(1, Program);
        w.WriteFixed32(2, Stream);
        w.WriteUInt32(3, LocalizationId);
        return w.ToArray();
    }
}

/// <summary>A presence change for one entity, whichever listener delivered it.</summary>
public sealed record BgsPresenceUpdate(BgsEntityKey EntityId, IReadOnlyList<BgsPresenceFieldOperation> Operations, bool IsFullState);

public sealed record BgsAccountPresence(
    BgsEntityKey Id,
    string? FullName,
    string? BattleTag,
    bool? IsOnline,
    string? BroadcastMessage,
    long? LastOnline,
    IReadOnlyList<BgsEntityKey> GameAccountIds);

public sealed record BgsGameAccountPresence(
    BgsEntityKey Id,
    bool? IsOnline,
    string? Program,
    uint? AwayStatus,
    bool? IsAfk,
    string? BattleTag,
    string? Name,
    BgsEntityKey? OwnerAccountId,
    long? LastOnline,
    BgsRichPresence? RichPresence,
    Variant? RichPresenceValue)
{
    public bool IsAway => IsAfk == true || ((AwayStatus ?? 0) & BgsPresenceFields.AwayFlag) != 0;

    public bool IsBusy => ((AwayStatus ?? 0) & BgsPresenceFields.BusyFlag) != 0;
}

/// <summary>What a friends list row needs: the same data SC:R's friends list gets.</summary>
public sealed record BgsFriendStatus(
    BgsEntityKey AccountId,
    string? BattleTag,
    string? FullName,
    bool IsOnline,
    string? Program,
    bool IsAway,
    bool IsBusy,
    BgsRichPresence? RichPresence,
    long? LastOnline,
    IReadOnlyList<BgsGameAccountPresence> GameAccounts);

/// <summary>
/// Holds the presence fields Battle.net has pushed, per entity, and reads the account and
/// game-account views out of them. Pure and not thread-safe; <see cref="BgsFriendsDirectory"/>
/// wraps it with a lock.
/// </summary>
public sealed class BgsPresenceTracker
{
    private readonly Dictionary<BgsEntityKey, Dictionary<BgsFieldKey, Variant>> _fields = [];
    private readonly HashSet<BgsEntityKey> _knownGameAccounts = [];

    /// <summary>
    /// Applies <paramref name="update"/> and returns the game-account ids it revealed for the first
    /// time (from an account's BN/1/3 list), so the caller can subscribe to their presence too:
    /// "online" and "program" live on the game accounts, not on the account.
    /// </summary>
    public IReadOnlyList<EntityId> Apply(BgsPresenceUpdate update)
    {
        if (!_fields.TryGetValue(update.EntityId, out var fields) || update.IsFullState)
        {
            fields = [];
            _fields[update.EntityId] = fields;
        }

        foreach (var op in update.Operations)
        {
            var key = BgsFieldKey.From(op.Key);
            if (op.Clear)
            {
                fields.Remove(key);
            }
            else
            {
                fields[key] = op.Value;
            }
        }

        List<EntityId> discovered = [];
        foreach (var gameAccount in GameAccountIdsOf(fields))
        {
            if (_knownGameAccounts.Add(gameAccount))
            {
                discovered.Add(gameAccount.ToEntityId());
            }
        }

        return discovered;
    }

    public IReadOnlyDictionary<BgsFieldKey, Variant> GetFields(BgsEntityKey entity) =>
        _fields.TryGetValue(entity, out var fields) ? new Dictionary<BgsFieldKey, Variant>(fields) : new Dictionary<BgsFieldKey, Variant>();

    public BgsAccountPresence? GetAccount(BgsEntityKey account)
    {
        if (!_fields.TryGetValue(account, out var fields))
        {
            return null;
        }

        var status = Get(fields, BgsPresenceFields.AccountGroup, BgsPresenceFields.AccountOnlineOrBroadcast);
        return new BgsAccountPresence(
            account,
            Get(fields, BgsPresenceFields.AccountGroup, BgsPresenceFields.AccountFullName)?.StringValue,
            Get(fields, BgsPresenceFields.AccountGroup, BgsPresenceFields.AccountBattleTag)?.StringValue,
            status?.BoolValue,
            status?.StringValue,
            AsInt(Get(fields, BgsPresenceFields.AccountGroup, BgsPresenceFields.AccountLastOnline)),
            GameAccountIdsOf(fields).ToList());
    }

    public BgsGameAccountPresence? GetGameAccount(BgsEntityKey gameAccount)
    {
        if (!_fields.TryGetValue(gameAccount, out var fields))
        {
            return null;
        }

        const uint g = BgsPresenceFields.GameAccountGroup;
        var online = Get(fields, g, BgsPresenceFields.GameAccountIsOnline);
        var away = AsInt(Get(fields, g, BgsPresenceFields.GameAccountAwayStatus));
        var afk = Get(fields, g, BgsPresenceFields.GameAccountAfk);
        var owner = AsEntity(Get(fields, g, BgsPresenceFields.GameAccountOwner));
        var rich = Get(fields, g, BgsPresenceFields.GameAccountRichPresence);
        return new BgsGameAccountPresence(
            gameAccount,
            online is null ? null : AsBool(online),
            AsFourCc(Get(fields, g, BgsPresenceFields.GameAccountProgram)),
            away is null ? null : (uint)away.Value,
            afk is null ? null : AsBool(afk),
            Get(fields, g, BgsPresenceFields.GameAccountBattleTag)?.StringValue,
            Get(fields, g, BgsPresenceFields.GameAccountName)?.StringValue,
            owner is null ? null : BgsEntityKey.From(owner),
            AsInt(Get(fields, g, BgsPresenceFields.GameAccountLastOnline)),
            BgsRichPresence.Decode(rich?.MessageValue ?? rich?.BlobValue),
            rich);
    }

    /// <summary>The account's game accounts: those its BN/1/3 list names plus any whose BN/2/7 owner field points back at it.</summary>
    public IReadOnlyList<BgsGameAccountPresence> GetGameAccountsOf(BgsEntityKey account)
    {
        var ids = new List<BgsEntityKey>();
        if (_fields.TryGetValue(account, out var fields))
        {
            ids.AddRange(GameAccountIdsOf(fields));
        }

        foreach (var (entity, entityFields) in _fields)
        {
            var owner = AsEntity(Get(entityFields, BgsPresenceFields.GameAccountGroup, BgsPresenceFields.GameAccountOwner));
            if (owner is not null && BgsEntityKey.From(owner) == account && !ids.Contains(entity))
            {
                ids.Add(entity);
            }
        }

        return ids.Select(GetGameAccount).OfType<BgsGameAccountPresence>().ToList();
    }

    private static IEnumerable<BgsEntityKey> GameAccountIdsOf(Dictionary<BgsFieldKey, Variant> fields)
    {
        foreach (var (key, value) in fields)
        {
            if (key.Program == BgsPresenceFields.BattleNetProgram
                && key.Group == BgsPresenceFields.AccountGroup
                && key.Field == BgsPresenceFields.AccountGameAccounts
                && AsEntity(value) is { } id)
            {
                yield return BgsEntityKey.From(id);
            }
        }
    }

    private static Variant? Get(Dictionary<BgsFieldKey, Variant> fields, uint group, uint field)
    {
        Variant? found = null;
        foreach (var (key, value) in fields)
        {
            if (key.Program == BgsPresenceFields.BattleNetProgram && key.Group == group && key.Field == field)
            {
                if (key.UniqueId == 0)
                {
                    return value;
                }

                found ??= value;
            }
        }

        return found;
    }

    private static bool AsBool(Variant v) => v.BoolValue ?? (v.IntValue ?? (long?)v.UintValue ?? 0) != 0;

    private static long? AsInt(Variant? v) => v?.IntValue ?? (long?)v?.UintValue;

    private static string? AsFourCc(Variant? v)
    {
        if (v is null)
        {
            return null;
        }

        if (!string.IsNullOrEmpty(v.FourccValue))
        {
            return v.FourccValue;
        }

        var number = v.UintValue ?? (ulong?)v.IntValue;
        return number is null ? v.StringValue : FourCc.Decode((uint)number.Value);
    }

    private static EntityId? AsEntity(Variant? v)
    {
        if (v is null)
        {
            return null;
        }

        if (v.EntityIdValue is not null)
        {
            return v.EntityIdValue;
        }

        // mooege packs EntityIds as a serialized message instead of entity_id_value.
        var bytes = v.MessageValue ?? v.BlobValue;
        if (bytes is null || bytes.Length == 0)
        {
            return null;
        }

        try
        {
            return EntityId.Decode(bytes);
        }
        catch (Exception ex) when (ex is IndexOutOfRangeException or ArgumentOutOfRangeException or InvalidOperationException)
        {
            return null;
        }
    }
}

/// <summary>
/// The Battle.net account friends list plus their presence, as one thread-safe model: feed it
/// FriendsService.Subscribe's response, FriendsListener callbacks and presence updates, read
/// <see cref="GetFriends"/>. <see cref="FrontClient"/> keeps one up to date in <see cref="FrontClient.Social"/>.
/// </summary>
public sealed class BgsFriendsDirectory
{
    private static readonly HashSet<string> LauncherPrograms = new(StringComparer.OrdinalIgnoreCase) { "BSAp", "App", "BN", "CLNT" };

    private readonly object _gate = new();
    private readonly Dictionary<BgsEntityKey, BgsFriendMessage> _friends = [];
    private readonly Dictionary<ulong, BgsReceivedInvitation> _received = [];
    private readonly BgsPresenceTracker _presence = new();

    /// <summary>Replaces the list with the one a FriendsService.Subscribe response carried.</summary>
    public void Apply(BgsFriendsSubscribeResponse response)
    {
        lock (_gate)
        {
            _friends.Clear();
            foreach (var f in response.Friends) _friends[BgsEntityKey.From(f.AccountId)] = f;
            _received.Clear();
            foreach (var i in response.ReceivedInvitations) _received[i.Id] = i;
        }
    }

    public void Apply(BgsFriendsNotification notification)
    {
        lock (_gate)
        {
            switch (notification.Kind)
            {
                case BgsFriendsNotificationKind.FriendAdded or BgsFriendsNotificationKind.FriendUpdated when notification.Friend is not null:
                    var key = BgsEntityKey.From(notification.Friend.AccountId);
                    _friends[key] = Merge(_friends.GetValueOrDefault(key), notification.Friend);
                    break;
                case BgsFriendsNotificationKind.FriendRemoved when notification.Friend is not null:
                    _friends.Remove(BgsEntityKey.From(notification.Friend.AccountId));
                    break;
                case BgsFriendsNotificationKind.ReceivedInvitationAdded when notification.ReceivedInvitation is not null:
                    _received[notification.ReceivedInvitation.Id] = notification.ReceivedInvitation;
                    break;
                case BgsFriendsNotificationKind.ReceivedInvitationRemoved when notification.InvitationId is not null:
                    _received.Remove(notification.InvitationId.Value);
                    break;
            }
        }
    }

    /// <inheritdoc cref="BgsPresenceTracker.Apply"/>
    public IReadOnlyList<EntityId> Apply(BgsPresenceUpdate update)
    {
        lock (_gate)
        {
            return _presence.Apply(update);
        }
    }

    public IReadOnlyList<EntityId> FriendAccountIds
    {
        get
        {
            lock (_gate)
            {
                return _friends.Keys.Select(k => k.ToEntityId()).ToList();
            }
        }
    }

    public IReadOnlyList<BgsReceivedInvitation> ReceivedInvitations
    {
        get
        {
            lock (_gate)
            {
                return _received.Values.ToList();
            }
        }
    }

    public IReadOnlyList<BgsFriendStatus> GetFriends()
    {
        lock (_gate)
        {
            return _friends.Values
                .Select(Describe)
                .OrderByDescending(f => f.IsOnline)
                .ThenBy(f => f.BattleTag ?? f.FullName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
    }

    public BgsFriendStatus? GetFriend(BgsEntityKey accountId)
    {
        lock (_gate)
        {
            return _friends.TryGetValue(accountId, out var friend) ? Describe(friend) : null;
        }
    }

    public IReadOnlyDictionary<BgsFieldKey, Variant> GetPresenceFields(BgsEntityKey entity)
    {
        lock (_gate)
        {
            return _presence.GetFields(entity);
        }
    }

    private BgsFriendStatus Describe(BgsFriendMessage friend)
    {
        var id = BgsEntityKey.From(friend.AccountId);
        var account = _presence.GetAccount(id);
        var games = _presence.GetGameAccountsOf(id);
        var online = games.Where(g => g.IsOnline == true).ToList();

        // A friend signed into both the Battle.net app and a game shows the game, as the app does.
        var primary = online
            .OrderBy(g => g.Program is not null && LauncherPrograms.Contains(g.Program) ? 1 : 0)
            .FirstOrDefault();
        var isOnline = online.Count > 0 || (games.Count == 0 && account?.IsOnline == true);

        return new BgsFriendStatus(
            id,
            friend.BattleTag ?? account?.BattleTag ?? games.Select(g => g.BattleTag).FirstOrDefault(t => t is not null),
            friend.FullName ?? account?.FullName,
            isOnline,
            primary?.Program,
            primary?.IsAway ?? false,
            primary?.IsBusy ?? false,
            primary?.RichPresence,
            account?.LastOnline ?? games.Select(g => g.LastOnline).Max(),
            games);
    }

    private static BgsFriendMessage Merge(BgsFriendMessage? old, BgsFriendMessage update) =>
        old is null
            ? update
            : new BgsFriendMessage
            {
                AccountId = update.AccountId,
                Attributes = update.Attributes.Count > 0 ? update.Attributes : old.Attributes,
                Roles = update.Roles.Count > 0 ? update.Roles : old.Roles,
                Privileges = update.Privileges ?? old.Privileges,
                AttributesEpoch = update.AttributesEpoch ?? old.AttributesEpoch,
                FullName = update.FullName ?? old.FullName,
                BattleTag = update.BattleTag ?? old.BattleTag,
                CreationTime = update.CreationTime ?? old.CreationTime,
            };
}
