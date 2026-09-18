using System.Runtime.CompilerServices;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>
/// Points every store at a scratch folder before the first test runs, so nothing a test does can
/// read or write a real user's settings — including stores whose tests never set an override of
/// their own (clan roster, tracked users, recent messages).
/// </summary>
internal static class TestConfigDirectory
{
#pragma warning disable CA2255 // A test assembly is exactly where a module initializer belongs.
    [ModuleInitializer]
#pragma warning restore CA2255
    internal static void RedirectToScratch()
    {
        var directory = Path.Combine(Path.GetTempPath(), "invigoration-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        ConfigStore.DirectoryOverride = directory;
        AppDomain.CurrentDomain.ProcessExit += (_, _) =>
        {
            try
            {
                Directory.Delete(directory, recursive: true);
            }
            catch (IOException)
            {
                // Left for the OS's own temp cleanup.
            }
            catch (UnauthorizedAccessException)
            {
            }
        };
    }
}

/// <summary>Guards the redirect itself — if it ever stopped applying, every store's tests would quietly go back to writing a real user's files.</summary>
public class TestConfigDirectoryTests
{
    [Fact]
    public void StoresUseAScratchFolder_NotTheRealOne()
    {
        var real = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Invigoration");

        Assert.NotEqual(real, ConfigStore.DefaultConfigDirectory());
        Assert.StartsWith(Path.GetTempPath(), ConfigStore.DefaultConfigDirectory());
        Assert.StartsWith(ConfigStore.DefaultConfigDirectory(), Invigoration.Core.Tracking.ProtocolUserTrackingStore.FilePath);
        Assert.StartsWith(ConfigStore.DefaultConfigDirectory(), Invigoration.Core.Clan.ClanRosterStore.FilePath);
        Assert.StartsWith(ConfigStore.DefaultConfigDirectory(), Invigoration.Core.Tracking.RecentMessageStore.FilePath);
    }
}
