namespace Invigoration.Core.Protocol;

/// <summary>
/// Battle.net Chat Server (BNCS) packet IDs, as documented at bnetdocs.org.
/// Only includes the subset this bot sends or handles.
/// </summary>
public enum BncsPacketId : byte
{
    SID_NULL = 0x00,
    SID_REPORTVERSION = 0x07,
    SID_ENTERCHAT = 0x0A,
    SID_GETCHANNELLIST = 0x0B,
    SID_JOINCHANNEL = 0x0C,
    SID_CHATCOMMAND = 0x0E,
    SID_CHATEVENT = 0x0F,
    SID_LEAVECHAT = 0x10,
    SID_UDPPINGRESPONSE = 0x14,
    SID_CHECKAD = 0x15,
    SID_MESSAGEBOX = 0x19,
    SID_PING = 0x25,
    SID_READUSERDATA = 0x26,
    SID_WRITEUSERDATA = 0x27,
    SID_CREATEACCOUNT = 0x2A,
    SID_GETICONDATA = 0x2D,
    SID_CHANGEPASSWORD = 0x31,
    SID_QUERYREALMS = 0x34,
    SID_LOGONRESPONSE2 = 0x3A,
    SID_LOGONREALMEX = 0x3E,
    SID_NEWS_INFO = 0x46,

    /// <summary>
    /// "Battle.net requests required work from the Client... ExtraWork has
    /// been used by Battle.net to collect statistics on system hardware and
    /// to prevent hacking/botting" (bnetdocs.org/document/43/extrawork).
    /// Compliance means downloading a server-provided MPQ, extracting a
    /// native DLL from it, and executing a function inside that DLL —
    /// deliberately never implemented here (see HandleRequiredWork).
    /// </summary>
    SID_REQUIREDWORK = 0x4C,
    SID_AUTH_INFO = 0x50,
    SID_AUTH_CHECK = 0x51,
    SID_AUTH_ACCOUNTCREATE = 0x52,
    SID_AUTH_ACCOUNTLOGON = 0x53,
    SID_AUTH_ACCOUNTLOGONPROOF = 0x54,
    SID_SETEMAIL = 0x59,
    SID_FRIENDSLIST = 0x65,
    SID_FRIENDSUPDATE = 0x66,
    SID_FRIENDSADD = 0x67,
    SID_FRIENDSREMOVE = 0x68,
    SID_FRIENDSPOSITION = 0x69,
    SID_CLANINFO = 0x75,
    SID_CLANINVITATIONRESPONSE = 0x79,
}
