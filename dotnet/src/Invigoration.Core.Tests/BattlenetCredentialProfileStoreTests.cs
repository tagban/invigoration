using System.Reflection;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

[Collection("BattlenetCredentialProfileStore")]
public class BattlenetCredentialProfileStoreTests
{
    [Fact]
    public void CreateAndSave_AssignsNonEmptyIdAndPersistsIt()
    {
        var name = $"profile-{Guid.NewGuid():N}";
        var profile = BattlenetCredentialProfileStore.CreateAndSave(name);
        try
        {
            Assert.False(string.IsNullOrEmpty(profile.Id));
            Assert.Equal(name, profile.Name);
            Assert.Contains(BattlenetCredentialProfileStore.Profiles, p => p.Id == profile.Id);
        }
        finally
        {
            BattlenetCredentialProfileStore.Profiles.RemoveAll(p => p.Id == profile.Id);
        }
    }

    [Fact]
    public void Find_UnknownId_ReturnsNull()
    {
        Assert.Null(BattlenetCredentialProfileStore.Find($"nonexistent-{Guid.NewGuid():N}"));
    }

    [Fact]
    public void Find_EmptyId_ReturnsNull()
    {
        Assert.Null(BattlenetCredentialProfileStore.Find(""));
    }

    [Fact]
    public void Renaming_DoesNotChangeCredentialFilePath()
    {
        var profile = BattlenetCredentialProfileStore.CreateAndSave("Original Name");
        try
        {
            var pathBefore = BattlenetCredentialProfileStore.CredentialFilePath(profile.Id);
            profile.Name = "Renamed";
            BattlenetCredentialProfileStore.Save();
            var pathAfter = BattlenetCredentialProfileStore.CredentialFilePath(profile.Id);

            Assert.Equal(pathBefore, pathAfter);
        }
        finally
        {
            BattlenetCredentialProfileStore.Profiles.RemoveAll(p => p.Id == profile.Id);
        }
    }

    [Fact]
    public void Delete_RemovesCachedCredentialFile()
    {
        var profile = BattlenetCredentialProfileStore.CreateAndSave("To Delete");
        var path = BattlenetCredentialProfileStore.CredentialFilePath(profile.Id);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, "fake-credential");

        BattlenetCredentialProfileStore.Delete(profile.Id);

        Assert.DoesNotContain(BattlenetCredentialProfileStore.Profiles, p => p.Id == profile.Id);
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void Delete_NoCredentialFileYet_DoesNotThrow()
    {
        var profile = BattlenetCredentialProfileStore.CreateAndSave("Never Signed In");

        BattlenetCredentialProfileStore.Delete(profile.Id);

        Assert.DoesNotContain(BattlenetCredentialProfileStore.Profiles, p => p.Id == profile.Id);
    }

    [Fact]
    public void HasCachedCredential_NoFile_IsFalse()
    {
        Assert.False(BattlenetCredentialProfileStore.HasCachedCredential($"nonexistent-{Guid.NewGuid():N}"));
    }

    [Fact]
    public void HasCachedCredential_ZeroLengthFile_IsFalse()
    {
        var id = Guid.NewGuid().ToString("N");
        var path = BattlenetCredentialProfileStore.CredentialFilePath(id);
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, "");

            Assert.False(BattlenetCredentialProfileStore.HasCachedCredential(id));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void HasCachedCredential_NonEmptyFile_IsTrue()
    {
        var id = Guid.NewGuid().ToString("N");
        var path = BattlenetCredentialProfileStore.CredentialFilePath(id);
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, "fake-credential");

            Assert.True(BattlenetCredentialProfileStore.HasCachedCredential(id));
        }
        finally
        {
            File.Delete(path);
        }
    }
}

[Collection("BattlenetCredentialProfileStore")]
public class EnsureBattlenetCredentialProfileIdTests
{
    private static string InvokeEnsure(BotEngine engine)
    {
        var method = typeof(BotEngine).GetMethod("EnsureBattlenetCredentialProfileId", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (string)method.Invoke(engine, [])!;
    }

    [Fact]
    public async Task EmptyProfileId_AutoCreatesAndStampsConfigAndFiresConfigPersistNeeded()
    {
        var config = new BotConfig { DisplayName = $"bot-{Guid.NewGuid():N}" };
        await using var engine = new BotEngine(config);
        var fired = 0;
        engine.ConfigPersistNeeded += () => fired++;

        try
        {
            var id = InvokeEnsure(engine);

            Assert.False(string.IsNullOrEmpty(id));
            Assert.Equal(id, config.BattlenetCredentialProfileId);
            Assert.Equal(1, fired);
            Assert.NotNull(BattlenetCredentialProfileStore.Find(id));
        }
        finally
        {
            BattlenetCredentialProfileStore.Profiles.RemoveAll(p => p.Id == config.BattlenetCredentialProfileId);
        }
    }

    [Fact]
    public async Task ValidExistingProfileId_ReturnsUnchangedAndDoesNotFireOrDuplicate()
    {
        var profile = BattlenetCredentialProfileStore.CreateAndSave("Existing");
        var config = new BotConfig { DisplayName = $"bot-{Guid.NewGuid():N}", BattlenetCredentialProfileId = profile.Id };
        await using var engine = new BotEngine(config);
        var fired = 0;
        engine.ConfigPersistNeeded += () => fired++;
        var countBefore = BattlenetCredentialProfileStore.Profiles.Count;

        try
        {
            var id = InvokeEnsure(engine);

            Assert.Equal(profile.Id, id);
            Assert.Equal(0, fired);
            Assert.Equal(countBefore, BattlenetCredentialProfileStore.Profiles.Count);
        }
        finally
        {
            BattlenetCredentialProfileStore.Profiles.RemoveAll(p => p.Id == profile.Id);
        }
    }

    [Fact]
    public async Task ProfileIdPointingAtDeletedProfile_Recreates()
    {
        var config = new BotConfig { DisplayName = $"bot-{Guid.NewGuid():N}", BattlenetCredentialProfileId = $"deleted-{Guid.NewGuid():N}" };
        await using var engine = new BotEngine(config);

        try
        {
            var id = InvokeEnsure(engine);

            Assert.NotEqual("", config.BattlenetCredentialProfileId);
            Assert.Equal(id, config.BattlenetCredentialProfileId);
            Assert.NotNull(BattlenetCredentialProfileStore.Find(id));
        }
        finally
        {
            BattlenetCredentialProfileStore.Profiles.RemoveAll(p => p.Id == config.BattlenetCredentialProfileId);
        }
    }
}
