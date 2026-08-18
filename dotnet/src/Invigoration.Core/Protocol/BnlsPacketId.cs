namespace Invigoration.Core.Protocol;

/// <summary>
/// Battle.net Login Server (BNLS) packet IDs, as documented at bnetdocs.org.
/// BNLS offloads Blizzard's CD-key/version/password hashing so the client
/// never has to implement CheckRevision or NLS/SRP itself.
/// </summary>
public enum BnlsPacketId : byte
{
    BNLS_NULL = 0x00,
    BNLS_CDKEY = 0x01,
    BNLS_LOGONCHALLENGE = 0x02,
    BNLS_LOGONPROOF = 0x03,
    BNLS_CREATEACCOUNT = 0x04,
    BNLS_VERSIONCHECK = 0x09,
    BNLS_HASHDATA = 0x0B,
    BNLS_CDKEY_EX = 0x0C,
    BNLS_CHOOSENLSREVISION = 0x0D,
    BNLS_AUTHORIZE = 0x0E,
    BNLS_AUTHORIZEPROOF = 0x0F,
    BNLS_REQUESTVERSIONBYTE = 0x10,
    BNLS_VERSIONCHECKEX2 = 0x1A,
}
