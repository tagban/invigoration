namespace Invigoration.Core.Auth;

/// <summary>
/// Mutable handshake state for one BNCS/BNLS session. Replaces the VB6
/// globals (GTC, CB, HType, SPass, CdkeyHash, hash(), AttemptedC, LRealm,
/// version, CheckSum, VerByte, Servers, statstring, cookie, versioncode) that
/// modBNET.bas/modBNLS.bas mutated directly — kept as one instance per
/// <see cref="Invigoration.Core.BotEngine"/> so multiple bots can run at once.
/// </summary>
public sealed class AuthState
{
    /// <summary>Client token, supplied by BNLS in its CD-key hash reply (was `GTC`).</summary>
    public uint ClientToken { get; set; }

    /// <summary>Server token, from the SID_AUTH_INFO reply (was `Servers`).</summary>
    public uint ServerToken { get; set; }

    /// <summary>Hashed CD-key blob from BNLS_CDKEY / BNLS_CDKEY_EX, forwarded verbatim into SID_AUTH_CHECK.</summary>
    public byte[] CdKeyHash { get; set; } = [];

    public uint ExeVersion { get; set; }
    public uint ExeChecksum { get; set; }
    public string ExeInfo { get; set; } = "";
    public uint VersionByte { get; set; }

    /// <summary>Which multi-step BNLS_HASHDATA flow is in progress (was `HType`).</summary>
    public HashPurpose HashPurpose { get; set; } = HashPurpose.None;

    /// <summary>Stage counter within the current hash flow (was `CB`).</summary>
    public int HashStage { get; set; }

    /// <summary>Double-hashed *old* password, staged during a change-password flow before the new password is hashed.</summary>
    public byte[] PendingOldPasswordDoubleHash { get; set; } = [];

    /// <summary>True once an account-creation attempt has been made this session, to avoid retry loops (was `AttemptedC`).</summary>
    public bool AttemptedAccountCreate { get; set; }

    /// <summary>True once logged into BNCS, informational only.</summary>
    public bool LoggedOnToBncs { get; set; }

    /// <summary>Set when D2/D2:LoD should continue into a realm (character server) logon after BNCS login (was `LRealm`).</summary>
    public bool WantsRealmLogon { get; set; }

    /// <summary>Set by a change-password command before the next SID_AUTH_CHECK success (was `Cpass`).</summary>
    public bool ChangePasswordRequested { get; set; }

    public string NewPassword { get; set; } = "";
}

/// <summary>Which outcome a BNLS_HASHDATA round-trip is working towards.</summary>
public enum HashPurpose
{
    None,

    /// <summary>Old login system: single-hash then double-hash for SID_LOGONRESPONSE2.</summary>
    AccountLogon,

    /// <summary>D2 realm (character server) logon: same double-hash flow as AccountLogon, but sends SID_LOGONREALMEX at the end.</summary>
    RealmLogon,

    /// <summary>Old account-creation system: single hash only, for SID_CREATEACCOUNT.</summary>
    AccountCreate,

    /// <summary>Change-password flow: double-hash the old password, single-hash the new one, for SID_CHANGEPASSWORD.</summary>
    ChangePassword,
}
