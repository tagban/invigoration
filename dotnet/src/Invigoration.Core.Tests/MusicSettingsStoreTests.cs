using Invigoration.Core.Music;

namespace Invigoration.Core.Tests;

public class MusicSettingsStoreTests
{
    [Fact]
    public void LoadsAnOlderFile_AndKeepsTheSpotifySignInUnreadableOnDisk()
    {
        var directory = Directory.CreateTempSubdirectory("invigoration-music-settings-").FullName;
        MusicSettingsStore.DirectoryOverride = directory;
        MusicSettingsStore.ResetCacheForTests();
        try
        {
            // What the YouTube Music player saved.
            File.WriteAllText(Path.Combine(directory, "music-settings.json"), """{"SelectedService":0,"IsEnabled":false,"ShowBottomBar":true}""");

            Assert.False(MusicSettingsStore.IsEnabled);
            Assert.True(MusicSettingsStore.ShowBottomBar);
            Assert.Equal("", MusicSettingsStore.SpotifyClientId);
            Assert.Equal("", MusicSettingsStore.SpotifyRefreshToken);
            Assert.False(MusicSettingsStore.SpotifyPromptAnswered);

            MusicSettingsStore.SpotifyClientId = "  my-client  ";
            MusicSettingsStore.SpotifyRefreshToken = "secret-refresh-token";
            MusicSettingsStore.SpotifyPromptAnswered = true;

            var onDisk = File.ReadAllText(Path.Combine(directory, "music-settings.json"));
            Assert.DoesNotContain("secret-refresh-token", onDisk);

            MusicSettingsStore.ResetCacheForTests();
            Assert.Equal("my-client", MusicSettingsStore.SpotifyClientId);
            Assert.Equal("secret-refresh-token", MusicSettingsStore.SpotifyRefreshToken);
            Assert.True(MusicSettingsStore.ShowBottomBar);
            Assert.True(MusicSettingsStore.SpotifyPromptAnswered);
        }
        finally
        {
            MusicSettingsStore.DirectoryOverride = null;
            MusicSettingsStore.ResetCacheForTests();
            Directory.Delete(directory, recursive: true);
        }
    }
}
