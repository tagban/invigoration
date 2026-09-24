namespace Invigoration.Scr;

/// <summary>
/// The retail StarCraft: Remastered client identity Battle.net expects, from build 1.23.10 as
/// recorded by ncarrillo/sc1-research (MIT). Battle.net checks these.
/// </summary>
public static class ScrProtocol
{
    public const string ProgramCode = "S1";
    public const uint Program = 0x5331;           // "S1"
    public const uint Platform = 0x4D633634;      // "Mc64"
    public const uint Locale = 0x656E5553;        // "enUS"
    public const uint SessionType = 0x44444354;   // "DDCT"
    public const string GameVersion = "1.23.10.13515";
    public const uint ApplicationVersion = 65559;
    public const uint ClientCapabilities = 0x00030100;
    public const string ClassicPath = "/S1/v2/rpc/client";

    /// <summary>The install identity the retail client sends: 20 bytes, as its base64 text.</summary>
    public const string ClientIdentity = "DJqHt+VTbDlhkzsfTvFlKrRHZjw=";

    // Classic services and methods (service hash, method hash).
    public const uint AuthenticationService = 0x17CDFF07;
    public const uint AuthSessionMethod = 0x95F59163;
    public const uint GameAccountService = 0x354252A4;
    public const uint GetToonsMethod = 0xBC18EDE5;

    /// <summary>
    /// GameAccount.CreateToon: {1 name, 2 uint64 gateway}, answered with {1 ToonInfo {1 id, 2 name, 3 gateway}}.
    /// The id and both layouts were read from the retail client's own library (libClientSdk).
    /// </summary>
    public const uint CreateToonMethod = 0x6697AC0A;
    public const uint GameVersionService = 0x3D930F0E;
    public const uint SetGameVersionMethod = 0xD48DE460;
    public const uint LegacyService = 0xD0C0F33D;
    public const uint LegacyConnectMethod = 0x607716CD;
    public const uint LegacyChatConnectMethod = 0x78D3F5A8;
    public const uint LegacyChatDisconnectMethod = 0x6DEB8B04;
    public const uint GatewayService = 0x2FD59FA3;
    public const uint GatewayUpdateMethod = 0xF5570066;

    /// <summary>A request trace in the SDK's shape, "RT-XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX", sent on the first classic call.</summary>
    public static byte[] NewRequestTrace()
    {
        var raw = System.Security.Cryptography.RandomNumberGenerator.GetBytes(16);
        var hex = Convert.ToHexString(raw);
        return System.Text.Encoding.ASCII.GetBytes($"RT-{hex[..8]}-{hex[8..12]}-{hex[12..16]}-{hex[16..20]}-{hex[20..]}");
    }
}
