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
    /// <summary>Client token — generated locally by CdKeyDecoder-based hashing, or supplied by BNLS's CD-key hash reply when that fallback is used (was `GTC`).</summary>
    public uint ClientToken { get; set; }

    /// <summary>Server token, from the SID_AUTH_INFO reply (was `Servers`).</summary>
    public uint ServerToken { get; set; }

    /// <summary>The CD-key block(s) sent verbatim in SID_AUTH_CHECK — either built locally by CdKeyDecoder or forwarded from BNLS_CDKEY/BNLS_CDKEY_EX.</summary>
    public byte[] CdKeyHash { get; set; } = [];

    /// <summary>The version check challenge from the last SID_AUTH_INFO reply, so BNLS's answer to it can be remembered (LogonCheckCache).</summary>
    public (Protocol.FileTimeValue FileTime, string FileName, string Formula)? VersionCheckChallenge { get; set; }

    /// <summary>True for a rapid-reconnect attempt that's logging on with LogonCheckCache's answers instead of asking BNLS.</summary>
    public bool UsingCachedChecks { get; set; }

    public uint ExeVersion { get; set; }
    public uint ExeChecksum { get; set; }
    public string ExeInfo { get; set; } = "";
    public uint VersionByte { get; set; }

    /// <summary>True once an account-creation attempt has been made this session, to avoid retry loops (was `AttemptedC`).</summary>
    public bool AttemptedAccountCreate { get; set; }

    /// <summary>True once logged into BNCS, informational only.</summary>
    public bool LoggedOnToBncs { get; set; }

    /// <summary>Set when D2/D2:LoD should continue into a realm (character server) logon after BNCS login (was `LRealm`).</summary>
    public bool WantsRealmLogon { get; set; }

    /// <summary>Set by a change-password command before the next SID_AUTH_CHECK success (was `Cpass`).</summary>
    public bool ChangePasswordRequested { get; set; }

    public string NewPassword { get; set; } = "";

    /// <summary>The exact password string the current logon hashed — the one typed, or its game-client casing after a retry (see BattlenetPassword). A D2 realm logon reuses whichever one the account logon succeeded with.</summary>
    public string LogonPassword { get; set; } = "";

    /// <summary>True once this logon has already retried with the password's casing changed, so a genuinely wrong password fails on the second rejection instead of looping.</summary>
    public bool RetriedPasswordCasing { get; set; }
}
