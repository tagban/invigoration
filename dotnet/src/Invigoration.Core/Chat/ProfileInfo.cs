namespace Invigoration.Core.Chat;

/// <summary>The classic BNCS "profile\*" fields for one account, as returned by SID_READUSERDATA (0x26). Age is displayed but not editable — see BotEngine.Profile.cs's WriteProfileAsync remarks.</summary>
public sealed record ProfileInfo(string Account, string Sex, string Age, string Location, string Description);
