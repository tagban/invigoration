namespace Invigoration.Core.Music.Pandora;

/// <summary>
/// The Android app's own partner identity — publicly known, not a secret we're extracting: every
/// open-source Pandora client (pydora, PianoBar, etc.) ships the same values, since Pandora's API
/// has no per-developer registration and this is simply which official client is "logging in".
/// Sourced from pydora's real code (<c>pandora/models/pandora.py</c>), confirmed there rather than
/// guessed from the docs page alone, which had gaps.
/// </summary>
public static class PandoraPartnerCredentials
{
    public const string Username = "android";
    public const string Password = "AC7IBG09A3DTSYM4R41UJWL07VLN8JI7";
    public const string DeviceModel = "android-generic";
    public const string DecryptKey = "R=U!LH$O2B#";
    public const string EncryptKey = "6#26FRL$ZWD";
}
