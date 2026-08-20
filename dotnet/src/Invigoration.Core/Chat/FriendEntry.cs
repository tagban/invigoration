namespace Invigoration.Core.Chat;

/// <summary>SID_FRIENDSLIST/ADD/UPDATE status bitflags, per bnetdocs.org.</summary>
[Flags]
public enum FriendStatus : byte
{
    None = 0,
    Mutual = 0x01,
    DoNotDisturb = 0x02,
    Away = 0x04,
}

/// <summary>Where a friend currently is, per bnetdocs.org's SID_FRIENDSLIST location id.</summary>
public enum FriendLocation : byte
{
    Offline = 0x00,
    NotInChat = 0x01,
    InChat = 0x02,
    PublicGame = 0x03,
    PrivateGame = 0x04,
    PrivateGameMutual = 0x05,
}

/// <summary>One entry in the classic BNCS friends list. <see cref="ProductCode"/> is the same wire-form 4-character product code used throughout this codebase (e.g. "VD2D"), suitable for <see cref="ChatIcon.GetProductIconKey"/>.</summary>
public sealed record FriendEntry(string Account, FriendStatus Status, FriendLocation Location, string ProductCode, string LocationName);

/// <summary>The fields SID_FRIENDSUPDATE carries — everything about a <see cref="FriendEntry"/> except the account name, which that packet identifies by list position instead.</summary>
public sealed record FriendStatusUpdate(FriendStatus Status, FriendLocation Location, string ProductCode, string LocationName);
