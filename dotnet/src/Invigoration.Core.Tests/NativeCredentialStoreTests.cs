using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>The native StarCraft II client's saved "keep me signed in" credentials, one per game.</summary>
[Collection("BattlenetCredentialProfileStore")]
public class NativeCredentialStoreTests
{
    [Fact]
    public void SavingReplacesTheOldCredential_AndLeavesNoTemporaryFile()
    {
        var profileId = $"native-{Guid.NewGuid():N}";

        BattlenetCredentialProfileStore.SaveNativeCredential(profileId, "S2", [1, 2, 3]);
        BattlenetCredentialProfileStore.SaveNativeCredential(profileId, "S2", [4, 5]);

        var path = BattlenetCredentialProfileStore.NativeCredentialFilePath(profileId, "S2");
        Assert.Equal([4, 5], BattlenetCredentialProfileStore.LoadNativeCredential(profileId, "S2"));
        Assert.False(File.Exists(path + ".new"));
        if (!OperatingSystem.IsWindows())
        {
            Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(path));
        }
    }

    [Fact]
    public void EachGameHasItsOwnCredential_AndAnEmptyOneIsNeverSaved()
    {
        var profileId = $"native-{Guid.NewGuid():N}";

        BattlenetCredentialProfileStore.SaveNativeCredential(profileId, "S2", [1]);
        BattlenetCredentialProfileStore.SaveNativeCredential(profileId, "S1", [2]);
        BattlenetCredentialProfileStore.SaveNativeCredential(profileId, "S2", []);

        Assert.Equal([1], BattlenetCredentialProfileStore.LoadNativeCredential(profileId, "S2"));
        Assert.Equal([2], BattlenetCredentialProfileStore.LoadNativeCredential(profileId, "S1"));
        Assert.Null(BattlenetCredentialProfileStore.LoadNativeCredential(profileId, "W3"));
    }

    [Fact]
    public void NativeCredentialsNeverShareStimpaksFile()
    {
        Assert.NotEqual(
            BattlenetCredentialProfileStore.CredentialFilePath("p"),
            BattlenetCredentialProfileStore.NativeCredentialFilePath("p", "S2"));
    }
}
