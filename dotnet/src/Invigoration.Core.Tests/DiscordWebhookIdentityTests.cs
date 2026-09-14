using System.Text.Json;
using Invigoration.Core.Discord;

namespace Invigoration.Core.Tests;

public class DiscordWebhookIdentityTests
{
    [Theory]
    [InlineData("RATS", "sc")]
    [InlineData("PXES", "scbw")]
    [InlineData("NB2W", "war2")]
    [InlineData("PX2DUSEast,Kilua,abc", "d2exp")]
    [InlineData("3RAW 5H3W 20", "war3")]
    [InlineData("PX3W", "w3tft")]
    [InlineData("TAHC", "chat")]
    public void AvatarUrlFor_PointsAtTheProductsClassicLogo(string product, string file)
    {
        Assert.Equal($"https://raw.githubusercontent.com/tagban/invigoration/main/docs/discord-avatars/{file}.png", DiscordWebhookIdentity.AvatarUrlFor(product));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("XYZ")]
    [InlineData("ZZZZ")]
    public void AvatarUrlFor_UnknownProduct_IsNull(string? product)
    {
        Assert.Null(DiscordWebhookIdentity.AvatarUrlFor(product));
    }

    // Every mapped logo must actually exist in docs/discord-avatars, or Discord shows a broken avatar.
    [Fact]
    public void EveryMappedAvatarFileExistsInTheRepo()
    {
        var repoRoot = AppContext.BaseDirectory;
        while (repoRoot is not null && !Directory.Exists(Path.Combine(repoRoot, "docs", "discord-avatars")))
        {
            repoRoot = Path.GetDirectoryName(repoRoot);
        }

        Assert.NotNull(repoRoot);
        foreach (var product in new[] { "RATS", "PXES", "RTSJ", "RHSS", "NB2W", "LTRD", "RHSD", "VD2D", "PX2D", "3RAW", "PX3W", "TAHC" })
        {
            var file = DiscordWebhookIdentity.AvatarUrlFor(product)![DiscordWebhookIdentity.AvatarBaseUrl.Length..];
            Assert.True(File.Exists(Path.Combine(repoRoot, "docs", "discord-avatars", file)), $"missing docs/discord-avatars/{file}");
        }
    }

    [Theory]
    [InlineData("Tagban", true)]
    [InlineData("Kilua*tagban", true)]
    [InlineData("", false)]
    [InlineData("   ", false)]
    [InlineData("DiscordFan", false)]
    [InlineData("clydebot", false)]
    public void IsUsableUsername_FollowsDiscordsRules(string name, bool expected)
    {
        Assert.Equal(expected, DiscordWebhookIdentity.IsUsableUsername(name));
        Assert.False(DiscordWebhookIdentity.IsUsableUsername(new string('a', 81)));
    }

    [Fact]
    public void BuildPayload_CarriesNameAvatarAndNeverAllowsMentions()
    {
        var json = DiscordWebhookIdentity.BuildPayload("@everyone hi", "Tagban", "https://example/sc.png");

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        Assert.Equal("@everyone hi", root.GetProperty("content").GetString());
        Assert.Equal("Tagban", root.GetProperty("username").GetString());
        Assert.Equal("https://example/sc.png", root.GetProperty("avatar_url").GetString());
        Assert.Equal(0, root.GetProperty("allowed_mentions").GetProperty("parse").GetArrayLength());
    }

    [Fact]
    public void BuildPayload_OmitsTheAvatarWhenThereIsNone()
    {
        using var doc = JsonDocument.Parse(DiscordWebhookIdentity.BuildPayload("hi", "Tagban", null));
        Assert.False(doc.RootElement.TryGetProperty("avatar_url", out _));
    }
}
