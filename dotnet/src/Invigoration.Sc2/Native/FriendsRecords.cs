namespace Invigoration.Sc2.Native;

/// <summary>Battlenet::Friends::PriorFriendId / FriendContainer5's identity — either an account or a specific character (toon).</summary>
public abstract record FriendIdentity
{
    private FriendIdentity()
    {
    }

    public sealed record Account(uint AccountId) : FriendIdentity;

    /// <summary>Battlenet::Toon::Handle, reached via FriendContainer5::Character or a Remove update's PriorFriendId::ToonHandle. Field order (program, region, realm, id) matches the already-established convention in <see cref="ToonRecordDecoder"/> and <see cref="PlayerTarget.ToonHandle"/>.</summary>
    public sealed record Character(uint ProgramId, byte Region, uint Realm, ulong Id) : FriendIdentity;
}

/// <summary>Battlenet::Friends::FriendContainer5's payload once resolved to a flat entry. Mirrors core/src/native/model.rs's FriendEntry — a friend can be identified by account or character, with the rest populated only for whichever variant produced it.</summary>
public sealed record FriendEntry(
    FriendIdentity Identity,
    string? DisplayName,
    string? FullName,
    string? Note,
    PlayerTarget.ProfileRecordAddress? Profile,
    ToonFullName? ToonName);

public enum SocialOperation
{
    Add,
    Remove,
    Modify,
}

/// <summary>One entry in a FriendsListNotify5 snapshot/delta. Mirrors core/src/native/model.rs's FriendUpdate.</summary>
public sealed record FriendUpdate(SocialOperation Operation, FriendEntry Entry);

/// <summary>Decoded payload of a FriendsListNotify5 record (Friends slot, command 30). Mirrors core/src/native/model.rs's FriendsPage.</summary>
public sealed record FriendsListRecord(IReadOnlyList<FriendUpdate> Updates, bool? Complete);

/// <summary>One toon belonging to a friend, as reported by a ToonsOfFriendsNotify record. Mirrors core/src/native/model.rs's FriendToon.</summary>
public sealed record FriendToon(uint AccountId, uint ProgramId, PlayerTarget.ProfileRecordAddress? Profile, ToonFullName ToonName);

/// <summary>Decoded payload of a ToonsOfFriendsNotify record (Friends slot, command 6) — sent in reply to a ToonsOfFriendsRequest for one friend's account. Mirrors core/src/native/model.rs's FriendToonPage.</summary>
public sealed record ToonsOfFriendsRecord(IReadOnlyList<FriendToon> Entries, bool Complete);
