namespace Invigoration.Sc2.Front;

/// <summary>
/// Builds the Front AuthenticationServer/1 LogonRequest. Field values are
/// exactly as documented at https://superioritybot.com/PROTOCOL's Front RPC
/// section for the "Base97563" client build: program/platform/locale are
/// fixed, application_version is always 0 regardless of the real
/// installation build number, and version must carry this specific embedded
/// SDK identity string, not a version number.
/// </summary>
public static class LogonBuilder
{
    public const string EmbeddedSdkVersion = "Battle.net Game Service SDK v1.48.2 \"cf68e241e0\"/104 (Jul 14 2026 19:45:54)";

    public static LogonRequest BuildLogonRequest(byte[]? cachedWebCredentials = null) => new()
    {
        Program = "S2",
        Platform = "Mc64",
        Locale = "enUS",
        Version = EmbeddedSdkVersion,
        ApplicationVersion = 0,
        AllowLogonQueueNotifications = true,
        CachedWebCredentials = cachedWebCredentials,
    };
}
