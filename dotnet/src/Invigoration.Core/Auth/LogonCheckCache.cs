using System.Collections.Concurrent;
using Invigoration.Core.Protocol;

namespace Invigoration.Core.Auth;

/// <summary>
/// What BNLS answered for the two parts of a classic logon only it can compute — a product's
/// version byte, and the version check for one server challenge — kept for the rest of the run.
/// A rapid reconnect (BotEngine.RapidReconnect.cs) reuses them to log straight back on without a
/// single BNLS round trip: the CD key and password are hashed locally already, so with these in
/// hand the whole logon is just the Battle.net server itself. Both answers depend only on what's
/// asked (the product; the challenge file and formula the server sent), never on which server or
/// bot asked, so every bot shares them.
/// </summary>
public static class LogonCheckCache
{
    /// <summary>A BNLS version check result, as SID_AUTH_CHECK sends it.</summary>
    public sealed record VersionCheck(uint ExeVersion, uint ExeChecksum, string ExeInfo);

    private static readonly ConcurrentDictionary<string, uint> VersionBytes = new(StringComparer.Ordinal);
    private static readonly ConcurrentDictionary<(string Product, FileTimeValue FileTime, string FileName, string Formula), VersionCheck> VersionChecks = new();

    public static void RememberVersionByte(string product, uint versionByte) => VersionBytes[product] = versionByte;

    public static bool TryGetVersionByte(string product, out uint versionByte) => VersionBytes.TryGetValue(product, out versionByte);

    public static void RememberVersionCheck(string product, FileTimeValue fileTime, string fileName, string formula, VersionCheck check) =>
        VersionChecks[(product, fileTime, fileName, formula)] = check;

    public static bool TryGetVersionCheck(string product, FileTimeValue fileTime, string fileName, string formula, out VersionCheck check) =>
        VersionChecks.TryGetValue((product, fileTime, fileName, formula), out check!);

    public static void ClearForTests()
    {
        VersionBytes.Clear();
        VersionChecks.Clear();
    }
}
