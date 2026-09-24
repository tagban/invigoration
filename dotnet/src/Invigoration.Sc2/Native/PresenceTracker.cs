using System.Buffers.Binary;

namespace Invigoration.Sc2.Native;

// Ported from ncarrillo/superiority (MIT): core/src/games/sc2/native/presence.rs's
// PresenceDirectory and core/src/games/sc2/chat/session.rs's friend_presence_id.

/// <summary>A friend's status as presence reports it.</summary>
public enum FriendPresence
{
    Online,
    Away,
    Busy,
    InGame,
    Offline,
}

/// <summary>
/// Tracks the presence records Battle.net pushes on Sunken (FieldSpecAnnounce and
/// PresenceUpdateNotify) and answers what state a friend is in. Presence names no
/// one directly: an update carries presence ids plus field values, and a presence
/// is tied to a friend through the values it holds: an account id field (0x1001b,
/// then 0x10005, then any field announced with type 21), a character handle field
/// (0x10019) or a profile address field (0x10014).
///
/// Not thread-safe; feed it from the receive loop and query it from there.
/// </summary>
public sealed class PresenceTracker
{
    public const uint FieldAvatar = 65_555;
    public const uint FieldToonProfile = 65_556;
    public const uint FieldToonHandle = 65_561;
    public const uint FieldAccountId = 0x0001_0005;
    public const uint FieldSocialAccountId = 0x0001_001b;
    public const uint FieldInGame = 0x50002;

    /// <summary>The SC2 clan tag: a count of byte pairs, a 0/1 odd-byte flag, then that many UTF-8 bytes. From ncarrillo/superiority (MIT).</summary>
    public const uint FieldClanTag = 0x50004;

    /// <summary>The game account a friend is on: region u8, program FourCC u32 (big-endian, e.g. "\0Fen", "BSAp", "\0\0S1"), id u32. Seen live 2026-09-24.</summary>
    public const uint FieldGameAccount = 0x10018;

    /// <summary>The account's BattleTag: u8 length, then UTF-8. Seen live 2026-09-24 on members and friends.</summary>
    public const uint FieldBattleTag = 0x1001E;

    /// <summary>The game account's name: region u8, program u32, u32, then a length (bytes minus 2) and UTF-8, e.g. an SC2 character "Crayfish#617" or a BattleTag.</summary>
    public const uint FieldGameAccountName = 0x1000E;
    public const uint FieldAway = 0x10020;
    public const uint FieldAwayFallback = 0x10010;
    public const uint FieldBusy = 0x10022;
    public const uint FieldBusyFallback = 0x10011;

    /// <summary>The field type id that carries an account id (from a capture of field 0x10015).</summary>
    public const byte TypeAccountInfo = 21;

    /// <summary>Fields a live presence holds. Once a presence has had one, losing both means it signed off.</summary>
    private static readonly uint[] SessionMarks = [0x0001_0003, 0x0001_0009];

    /// <summary>Away/busy fields. The one written most recently decides (ties go to the later entry).</summary>
    private static readonly (uint Handle, FriendPresence State)[] StandingOrder =
    [
        (FieldBusyFallback, FriendPresence.Busy),
        (FieldAwayFallback, FriendPresence.Away),
        (FieldBusy, FriendPresence.Busy),
        (FieldAway, FriendPresence.Away),
    ];

    private static readonly SortedDictionary<uint, byte[]> NoValues = [];
    private static readonly Dictionary<uint, ulong> NoStanding = [];

    private readonly Dictionary<uint, PresenceFieldSpec> _fields = [];
    private readonly SortedDictionary<uint, SortedDictionary<uint, byte[]>> _values = [];
    private readonly Dictionary<uint, uint> _aliases = [];
    private readonly Dictionary<uint, uint> _localPresence = [];
    private readonly Dictionary<uint, uint> _accountPresence = [];
    private readonly Dictionary<(uint Label, ulong Id), uint> _profilePresence = [];
    private readonly Dictionary<ToonKey, uint> _toonPresence = [];
    private readonly Dictionary<uint, bool> _online = [];
    private readonly Dictionary<uint, ulong> _seen = [];
    private readonly Dictionary<uint, Dictionary<uint, ulong>> _standing = [];
    private readonly HashSet<uint> _sessioned = [];
    private ulong _clock;

    /// <summary>Records the field definitions from a FieldSpecAnnounce (Presence 1). Later announcements add to or replace earlier ones.</summary>
    public void Announce(PresenceFieldsRecord record)
    {
        foreach (var field in record.Fields)
        {
            _fields[field.Handle] = field;
        }
    }

    /// <summary>
    /// Applies a PresenceUpdateNotify (Presence 0). Returns false, changing nothing,
    /// when the update is malformed: no presence id, or field sizes that don't match
    /// the data. Values after the first field that hasn't been announced are dropped,
    /// as upstream does.
    /// </summary>
    public bool Apply(PresenceUpdateRecord update)
    {
        var canonical = CanonicalId(update.LocalPresenceId, update.MasterPresenceId);
        if (canonical == 0)
        {
            return false;
        }

        var offset = 0;
        var variableIndex = 0;
        var unread = false;
        var decoded = new List<(uint Handle, byte[] Value)>(update.Handles.Count);
        foreach (var handle in update.Handles)
        {
            if (!_fields.TryGetValue(handle, out var field))
            {
                unread = true;
                break;
            }

            int size;
            if (field.FixedSize is { } fixedSize)
            {
                size = fixedSize;
            }
            else
            {
                if (variableIndex >= update.VariableSizes.Count)
                {
                    return false;
                }

                size = update.VariableSizes[variableIndex++];
            }

            if (offset + size > update.FieldData.Length)
            {
                return false;
            }

            decoded.Add((handle, update.FieldData.AsSpan(offset, size).ToArray()));
            offset += size;
        }

        if (!unread && (offset != update.FieldData.Length || variableIndex != update.VariableSizes.Count))
        {
            return false;
        }

        MergeAlias(update.LocalPresenceId, canonical);
        MergeAlias(update.MasterPresenceId, canonical);
        foreach (var id in (uint[])[update.LocalPresenceId, update.MasterPresenceId])
        {
            if (id != 0)
            {
                _aliases[id] = canonical;
            }
        }

        if (update.LocalPresenceId != 0)
        {
            _localPresence[canonical] = update.LocalPresenceId;
        }

        _clock++;
        var clock = _clock;
        var values = GetOrAdd(_values, canonical);
        foreach (var handle in update.ClearedHandles)
        {
            values.Remove(handle);
        }

        foreach (var (handle, value) in decoded)
        {
            values[handle] = value;
        }

        var standing = GetOrAdd(_standing, canonical);
        foreach (var handle in update.ClearedHandles)
        {
            standing.Remove(handle);
        }

        foreach (var (handle, _) in decoded)
        {
            if (IsStanding(handle))
            {
                standing[handle] = clock;
            }
        }

        var held = SessionMarks.Count(values.ContainsKey);
        if (held > 0)
        {
            _sessioned.Add(canonical);
        }

        var gone = held == 0 && _sessioned.Contains(canonical);
        _online[canonical] = update.Online && !gone;
        _seen[canonical] = clock;
        RebuildIdentityIndexes();
        return true;
    }

    /// <summary>
    /// The state of a friend identified by account or character, or null when no
    /// presence has been seen for them. Upstream shows a friend with no presence as
    /// offline.
    /// </summary>
    public FriendPresence? For(FriendIdentity identity) =>
        PresenceIdFor(identity) is { } presenceId ? State(presenceId) : null;

    /// <summary>
    /// As <see cref="For(FriendIdentity)"/>, but falls back to the friend's profile
    /// address when their account or character isn't linked to a presence.
    /// </summary>
    public FriendPresence? For(FriendEntry friend) =>
        PresenceIdFor(friend) is { } presenceId ? State(presenceId) : null;

    /// <summary>The friend's local presence id (a whisper can target it), or null when none is known.</summary>
    public uint? PresenceIdFor(FriendIdentity identity) => identity switch
    {
        FriendIdentity.Account account => _accountPresence.TryGetValue(account.AccountId, out var id) ? id : null,
        FriendIdentity.Character character =>
            _toonPresence.TryGetValue(new ToonKey(character.Region, character.ProgramId, character.Realm, character.Id), out var id) ? id : null,
        _ => null,
    };

    /// <summary>As <see cref="PresenceIdFor(FriendIdentity)"/>, falling back to the friend's profile address.</summary>
    public uint? PresenceIdFor(FriendEntry friend)
    {
        if (PresenceIdFor(friend.Identity) is { } id)
        {
            return id;
        }

        return friend.Profile is { } profile && _profilePresence.TryGetValue((profile.Label, profile.RecordId), out var byProfile)
            ? byProfile
            : null;
    }

    /// <summary>
    /// The portrait a presence advertises directly (field 0x10013: table then offset,
    /// 16 bits each, big-endian), or null when it hasn't sent one. When set, it wins over
    /// anything a profile read of <see cref="ProfileFor(uint)"/> would find.
    /// <paramref name="presenceId"/> may be a local or master id, e.g. a chat member's
    /// <see cref="MembershipChange.Join.PresenceId"/>.
    /// </summary>
    public Sc2Portrait? AvatarFor(uint presenceId) =>
        Value(presenceId, FieldAvatar) is { Length: >= 4 } value
            ? new Sc2Portrait(BinaryPrimitives.ReadUInt16BigEndian(value), BinaryPrimitives.ReadUInt16BigEndian(value.AsSpan(2)))
            : null;

    /// <summary>
    /// The profile record address a presence carries (field 0x10014: label u32 then id
    /// u64, big-endian), or null. Read it with a profile read to find the portrait when
    /// <see cref="AvatarFor(uint)"/> is null.
    /// </summary>
    public PlayerTarget.ProfileRecordAddress? ProfileFor(uint presenceId) =>
        Value(presenceId, FieldToonProfile) is { } value && DecodeProfileAddress(value) is { } address
            ? new PlayerTarget.ProfileRecordAddress(address.Label, address.Id)
            : null;

    /// <summary>As <see cref="AvatarFor(uint)"/>, for the presence linked to a friend's account or character.</summary>
    public Sc2Portrait? AvatarFor(FriendIdentity identity) =>
        PresenceIdFor(identity) is { } presenceId ? AvatarFor(presenceId) : null;

    /// <summary>As <see cref="AvatarFor(uint)"/>, for the presence linked to a friend (falling back to their profile address).</summary>
    public Sc2Portrait? AvatarFor(FriendEntry friend) =>
        PresenceIdFor(friend) is { } presenceId ? AvatarFor(presenceId) : null;

    /// <summary>A friend's profile address: the one the friends list gave, else the one their presence carries.</summary>
    public PlayerTarget.ProfileRecordAddress? ProfileFor(FriendEntry friend) =>
        friend.Profile ?? (PresenceIdFor(friend) is { } presenceId ? ProfileFor(presenceId) : null);

    /// <summary>The program the presence's game account is in ("S2", "S1", "Fen", "BSAp"...), from field 0x10018; null if not sent.</summary>
    public string? ProgramFor(uint presenceId) =>
        Value(presenceId, FieldGameAccount) is { Length: >= 5 } value
            ? System.Text.Encoding.ASCII.GetString(value, 1, 4).TrimStart('\0') is { Length: > 0 } program ? program : null
            : null;

    /// <summary>The account's BattleTag (field 0x1001E), or null.</summary>
    public string? BattleTagFor(uint presenceId) =>
        Value(presenceId, FieldBattleTag) is { Length: >= 2 } value && value[0] > 0 && value[0] <= value.Length - 1
            ? System.Text.Encoding.UTF8.GetString(value, 1, value[0])
            : null;

    /// <summary>The game account's name (field 0x1000E): for SC2 the character and its code, e.g. "Crayfish#617"; or null.</summary>
    public string? GameAccountNameFor(uint presenceId) =>
        Value(presenceId, FieldGameAccountName) is { Length: > 10 } value && value[9] + 2 <= value.Length - 10
            ? System.Text.Encoding.UTF8.GetString(value, 10, value[9] + 2)
            : null;

    /// <summary>The member's SC2 clan tag (field 0x50004), without brackets, or null.</summary>
    public string? ClanTagFor(uint presenceId)
    {
        if (Value(presenceId, FieldClanTag) is not { Length: >= 2 } value || value[1] > 1)
        {
            return null;
        }

        var length = value[0] * 2 + value[1];
        if (value.Length - 2 != length)
        {
            return null;
        }

        try
        {
            var tag = new System.Text.UTF8Encoding(false, true).GetString(value, 2, length).Trim();
            return tag.Length > 0 ? tag : null;
        }
        catch (System.Text.DecoderFallbackException)
        {
            return null;
        }
    }

    public string? ProgramFor(FriendEntry friend) => PresenceIdFor(friend) is { } presenceId ? ProgramFor(presenceId) : null;

    public string? BattleTagFor(FriendEntry friend) => PresenceIdFor(friend) is { } presenceId ? BattleTagFor(presenceId) : null;

    private byte[]? Value(uint presenceId, uint handle) =>
        _values.TryGetValue(CanonicalPresenceId(presenceId), out var values) && values.TryGetValue(handle, out var value)
            ? value
            : null;

    /// <summary>The state of one presence id, or null when nothing says whether it's online.</summary>
    public FriendPresence? State(uint presenceId)
    {
        var canonical = CanonicalPresenceId(presenceId);
        bool? online = _online.TryGetValue(canonical, out var isOnline) ? isOnline : null;
        return _values.TryGetValue(canonical, out var values)
            ? StateFromValues(values, online, _standing.GetValueOrDefault(canonical) ?? NoStanding)
            : StateFromValues(NoValues, online, NoStanding);
    }

    private static bool IsStanding(uint handle) =>
        handle == FieldInGame || StandingOrder.Any(entry => entry.Handle == handle);

    private static FriendPresence? StateFromValues(SortedDictionary<uint, byte[]> values, bool? online, Dictionary<uint, ulong> written)
    {
        if (online == false)
        {
            return FriendPresence.Offline;
        }

        if (BoolField(values, FieldInGame))
        {
            return FriendPresence.InGame;
        }

        (uint Handle, FriendPresence State)? latest = null;
        ulong latestWritten = 0;
        foreach (var entry in StandingOrder)
        {
            if (!values.ContainsKey(entry.Handle))
            {
                continue;
            }

            var when = written.GetValueOrDefault(entry.Handle);
            if (latest is null || when >= latestWritten)
            {
                latest = entry;
                latestWritten = when;
            }
        }

        if (latest is { } standing && BoolField(values, standing.Handle))
        {
            return standing.State;
        }

        return online switch
        {
            true => FriendPresence.Online,
            false => FriendPresence.Offline,
            null => null,
        };
    }

    private static bool BoolField(SortedDictionary<uint, byte[]> values, uint handle) =>
        values.TryGetValue(handle, out var value) && value is [1];

    private uint? AccountIdFrom(SortedDictionary<uint, byte[]> values)
    {
        foreach (var handle in (uint[])[FieldSocialAccountId, FieldAccountId])
        {
            if (values.TryGetValue(handle, out var value) && DecodeAccountId(value) is { } accountId)
            {
                return accountId;
            }
        }

        foreach (var (handle, value) in values)
        {
            if (_fields.TryGetValue(handle, out var field) && field.TypeId == TypeAccountInfo && DecodeAccountId(value) is { } accountId)
            {
                return accountId;
            }
        }

        return null;
    }

    private static uint? DecodeAccountId(byte[] value)
    {
        if (value.Length < 4)
        {
            return null;
        }

        var accountId = BinaryPrimitives.ReadUInt32BigEndian(value);
        return accountId != 0 ? accountId : null;
    }

    private static (uint Label, ulong Id)? DecodeProfileAddress(byte[] value) =>
        value.Length < 12
            ? null
            : (BinaryPrimitives.ReadUInt32BigEndian(value), BinaryPrimitives.ReadUInt64BigEndian(value.AsSpan(4)));

    private static ToonKey? DecodeToonHandle(byte[] value) =>
        value.Length < 17
            ? null
            : new ToonKey(
                value[0],
                BinaryPrimitives.ReadUInt32BigEndian(value.AsSpan(1)),
                BinaryPrimitives.ReadUInt32BigEndian(value.AsSpan(5)),
                BinaryPrimitives.ReadUInt64BigEndian(value.AsSpan(9)));

    private static int AccountFieldPriority(SortedDictionary<uint, byte[]> values, Dictionary<uint, PresenceFieldSpec> fields)
    {
        if (values.ContainsKey(FieldSocialAccountId))
        {
            return 2;
        }

        return values.ContainsKey(FieldAccountId)
            || values.Keys.Any(handle => fields.TryGetValue(handle, out var field) && field.TypeId == TypeAccountInfo)
            ? 1
            : 0;
    }

    private void RebuildIdentityIndexes()
    {
        // Folding merges presences; bounded so a fold that makes no progress can't loop forever.
        for (var pass = 0; pass < 16; pass++)
        {
            if (!TryRebuildIdentityIndexes())
            {
                return;
            }
        }
    }

    /// <summary>Rebuilds the account/profile/character indexes. Returns true when duplicate presences for one account were folded together and another pass is needed.</summary>
    private bool TryRebuildIdentityIndexes()
    {
        _accountPresence.Clear();
        _profilePresence.Clear();
        _toonPresence.Clear();
        var accountCandidates = new SortedDictionary<uint, (int Priority, uint PresenceId)>();
        var owners = new Dictionary<uint, uint>();
        var duplicates = new List<(uint Folded, uint Keep)>();

        foreach (var (canonical, values) in _values)
        {
            if (!_localPresence.TryGetValue(canonical, out var localPresenceId))
            {
                continue;
            }

            if (AccountIdFrom(values) is { } accountId)
            {
                if (owners.TryGetValue(accountId, out var held))
                {
                    var (keep, folded) = _seen.GetValueOrDefault(canonical) >= _seen.GetValueOrDefault(held)
                        ? (canonical, held)
                        : (held, canonical);
                    owners[accountId] = keep;
                    duplicates.Add((folded, keep));
                }
                else
                {
                    owners[accountId] = canonical;
                }

                var priority = AccountFieldPriority(values, _fields);
                if (!accountCandidates.TryGetValue(accountId, out var candidate) || priority > candidate.Priority)
                {
                    accountCandidates[accountId] = (priority, localPresenceId);
                }
            }

            if (values.TryGetValue(FieldToonProfile, out var profileValue) && DecodeProfileAddress(profileValue) is { } profile)
            {
                _profilePresence[profile] = localPresenceId;
            }

            if (values.TryGetValue(FieldToonHandle, out var toonValue) && DecodeToonHandle(toonValue) is { } toon)
            {
                _toonPresence[toon] = localPresenceId;
            }
        }

        var folds = 0;
        foreach (var (folded, keep) in duplicates)
        {
            var kept = CanonicalPresenceId(keep);
            if (folded != kept && _values.ContainsKey(folded))
            {
                FoldIdentity(folded, kept);
                folds++;
            }
        }

        if (folds > 0)
        {
            return true;
        }

        foreach (var (accountId, (_, presenceId)) in accountCandidates)
        {
            _accountPresence[accountId] = presenceId;
        }

        return false;
    }

    private void FoldIdentity(uint folded, uint keep)
    {
        if (_values.Remove(folded, out var values))
        {
            var kept = GetOrAdd(_values, keep);
            foreach (var (handle, value) in values)
            {
                kept.TryAdd(handle, value);
            }
        }

        if (_standing.Remove(folded, out var standing))
        {
            var kept = GetOrAdd(_standing, keep);
            foreach (var (handle, written) in standing)
            {
                kept.TryAdd(handle, written);
            }
        }

        if (_online.Remove(folded, out var online))
        {
            _online.TryAdd(keep, online);
        }

        if (_localPresence.Remove(folded, out var local))
        {
            _localPresence.TryAdd(keep, local);
        }

        if (_seen.Remove(folded, out var seen))
        {
            _seen[keep] = Math.Max(_seen.GetValueOrDefault(keep, seen), seen);
        }

        if (_sessioned.Remove(folded))
        {
            _sessioned.Add(keep);
        }

        RemapAliases(folded, keep);
        _aliases[folded] = keep;
    }

    private uint CanonicalId(uint local, uint master)
    {
        foreach (var id in (uint[])[master, local])
        {
            if (id != 0 && _aliases.TryGetValue(id, out var canonical))
            {
                return canonical;
            }
        }

        return master != 0 ? master : local;
    }

    private void MergeAlias(uint id, uint canonical)
    {
        if (!_aliases.TryGetValue(id, out var previous) || previous == canonical)
        {
            return;
        }

        if (_values.Remove(previous, out var previousValues))
        {
            var values = GetOrAdd(_values, canonical);
            foreach (var (handle, value) in previousValues)
            {
                values.TryAdd(handle, value);
            }
        }

        if (_online.Remove(previous, out var previousOnline))
        {
            _online.TryAdd(canonical, previousOnline);
        }

        if (_localPresence.Remove(previous, out var previousLocal))
        {
            _localPresence.TryAdd(canonical, previousLocal);
        }

        RemapAliases(previous, canonical);
    }

    private void RemapAliases(uint from, uint to)
    {
        foreach (var alias in _aliases.Where(pair => pair.Value == from).Select(pair => pair.Key).ToList())
        {
            _aliases[alias] = to;
        }
    }

    private uint CanonicalPresenceId(uint presenceId) =>
        _aliases.TryGetValue(presenceId, out var canonical) ? canonical : presenceId;

    private static TValue GetOrAdd<TValue>(IDictionary<uint, TValue> map, uint key)
        where TValue : new()
    {
        if (!map.TryGetValue(key, out var value))
        {
            value = new TValue();
            map[key] = value;
        }

        return value;
    }

    private readonly record struct ToonKey(byte Region, uint ProgramId, uint Realm, ulong Id);
}
