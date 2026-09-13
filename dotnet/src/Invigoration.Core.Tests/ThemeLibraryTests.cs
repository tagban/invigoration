using Invigoration.Core.Chat;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>Serialized: ThemeLibrary is static, with a process-wide cache and an on-disk folder.</summary>
[Collection("ThemeLibrary")]
public class ThemeLibraryTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "invig-themes-" + Guid.NewGuid().ToString("N"));

    public ThemeLibraryTests()
    {
        ThemeLibrary.DirectoryOverride = _dir;
        ThemeLibrary.ResetCacheForTests();
    }

    public void Dispose()
    {
        ThemeLibrary.DirectoryOverride = null;
        ThemeLibrary.ResetCacheForTests();
        try
        {
            Directory.Delete(_dir, recursive: true);
        }
        catch (DirectoryNotFoundException)
        {
        }
    }

    [Fact]
    public void BuiltIns_AreDefaultDiabloStarCraftWarcraft_InThatOrder()
    {
        Assert.Equal(
            [ThemeLibrary.DefaultId, ThemeLibrary.DiabloIIId, ThemeLibrary.StarCraftId, ThemeLibrary.WarcraftId],
            ThemeLibrary.All().Select(t => t.Id));
        Assert.All(ThemeLibrary.BuiltIns, t => Assert.True(t.IsBuiltIn));
    }

    [Fact]
    public void BuiltIns_EachDrawTheirOwnFrame()
    {
        Assert.Equal(ThemeFrameStyle.None, ThemeLibrary.Resolve(ThemeLibrary.DefaultId).Frame);
        Assert.Equal(ThemeFrameStyle.DiabloII, ThemeLibrary.Resolve(ThemeLibrary.DiabloIIId).Frame);
        Assert.Equal(ThemeFrameStyle.StarCraft, ThemeLibrary.Resolve(ThemeLibrary.StarCraftId).Frame);
        Assert.Equal(ThemeFrameStyle.Warcraft, ThemeLibrary.Resolve(ThemeLibrary.WarcraftId).Frame);
        Assert.Equal(ThemeLayout.CharacterDock, ThemeLibrary.Resolve(ThemeLibrary.DiabloIIId).Layout);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("custom-deleted-long-ago")]
    public void Resolve_UnknownId_FallsBackToDefault(string? id)
    {
        Assert.Equal(ThemeLibrary.DefaultId, ThemeLibrary.Resolve(id).Id);
    }

    // Upgrade path: configs saved before themes existed.
    [Theory]
    [InlineData("", false, ThemeLibrary.DefaultId)]
    [InlineData("", true, ThemeLibrary.DiabloIIId)]
    [InlineData(ThemeLibrary.WarcraftId, true, ThemeLibrary.WarcraftId)]
    [InlineData(ThemeLibrary.StarCraftId, false, ThemeLibrary.StarCraftId)]
    public void ThemeIdFor_MigratesTheOldD2StyleSwitch_ButAnExplicitThemeWins(string themeId, bool useD2, string expected)
    {
        Assert.Equal(expected, ThemeLibrary.ThemeIdFor(new BotConfig { ThemeId = themeId, UseD2ChatLayout = useD2 }));
    }

    [Fact]
    public void AssignTo_KeepsTheLegacyD2FlagInStep()
    {
        var config = new BotConfig();

        ThemeLibrary.AssignTo(config, ThemeLibrary.Resolve(ThemeLibrary.DiabloIIId));
        Assert.Equal(ThemeLibrary.DiabloIIId, config.ThemeId);
        Assert.True(config.UseD2ChatLayout);

        ThemeLibrary.AssignTo(config, ThemeLibrary.Resolve(ThemeLibrary.StarCraftId));
        Assert.False(config.UseD2ChatLayout);
    }

    [Fact]
    public void Duplicate_MakesAnIndependentCustomCopy()
    {
        var source = ThemeLibrary.Resolve(ThemeLibrary.StarCraftId);

        var copy = ThemeLibrary.Duplicate(source, "My StarCraft");
        copy.Materials.Accent = 0x00FF00;

        Assert.False(copy.IsBuiltIn);
        Assert.StartsWith("custom-", copy.Id);
        Assert.Equal(source.Frame, copy.Frame);
        Assert.NotEqual(0x00FF00, source.Materials.Accent);
    }

    [Fact]
    public void Save_ThenReload_RoundTripsEveryField()
    {
        var theme = ThemeLibrary.Duplicate(ThemeLibrary.Resolve(ThemeLibrary.DiabloIIId), "Hellforge");
        theme.Layout = ThemeLayout.Standard;
        theme.ShowChatGem = false;
        theme.ColorScheme = null;
        theme.HeaderFont = "Papyrus";
        theme.Materials.Trim = 0x123456;

        ThemeLibrary.Save(theme);
        ThemeLibrary.ResetCacheForTests();

        var loaded = ThemeLibrary.Resolve(theme.Id);
        Assert.Equal("Hellforge", loaded.Name);
        Assert.False(loaded.IsBuiltIn);
        Assert.Equal(ThemeFrameStyle.DiabloII, loaded.Frame);
        Assert.Equal(ThemeLayout.Standard, loaded.Layout);
        Assert.False(loaded.ShowChatGem);
        Assert.Null(loaded.ColorScheme);
        Assert.Equal("Papyrus", loaded.HeaderFont);
        Assert.Equal(0x123456, loaded.Materials.Trim);
        Assert.Contains(ThemeLibrary.All(), t => t.Id == theme.Id);
    }

    [Fact]
    public void BuiltIns_CantBeSavedOverOrDeleted()
    {
        var builtIn = ThemeLibrary.Resolve(ThemeLibrary.WarcraftId);

        Assert.Throws<InvalidOperationException>(() => ThemeLibrary.Save(builtIn));

        // Even a custom theme that claims a built-in's id.
        var impostor = ThemeLibrary.Duplicate(builtIn, "Impostor");
        impostor.Id = ThemeLibrary.WarcraftId;
        Assert.Throws<InvalidOperationException>(() => ThemeLibrary.Save(impostor));

        ThemeLibrary.Delete(ThemeLibrary.WarcraftId);
        Assert.Equal(ThemeLibrary.WarcraftId, ThemeLibrary.Resolve(ThemeLibrary.WarcraftId).Id);
    }

    [Fact]
    public void Delete_BotsOnThatThemeFallBackToDefault_AndListenersHear()
    {
        var theme = ThemeLibrary.Duplicate(ThemeLibrary.Resolve(ThemeLibrary.StarCraftId), "Short-lived");
        ThemeLibrary.Save(theme);
        var changes = 0;
        void OnChanged() => changes++;
        ThemeLibrary.ThemesChanged += OnChanged;
        try
        {
            ThemeLibrary.Delete(theme.Id);
        }
        finally
        {
            ThemeLibrary.ThemesChanged -= OnChanged;
        }

        Assert.Equal(1, changes);
        Assert.Equal(ThemeLibrary.DefaultId, ThemeLibrary.ResolveFor(new BotConfig { ThemeId = theme.Id }).Id);
    }

    [Fact]
    public void AMalformedThemeFile_IsSkipped_NotFatal()
    {
        var good = ThemeLibrary.Duplicate(ThemeLibrary.Resolve(ThemeLibrary.DiabloIIId), "Good");
        ThemeLibrary.Save(good);
        File.WriteAllText(Path.Combine(_dir, "broken.json"), "{ not json");
        ThemeLibrary.ResetCacheForTests();

        Assert.Contains(ThemeLibrary.All(), t => t.Id == good.Id);
    }

    [Fact]
    public void WarcraftTheme_HasItsOwnPalette()
    {
        Assert.Equal(ChatColorScheme.Warcraft, ThemeLibrary.Resolve(ThemeLibrary.WarcraftId).ColorScheme);
        Assert.Same(ChatPalette.Warcraft, ChatPalette.ForScheme(new BotConfig { ChatColorScheme = ChatColorScheme.Warcraft }));
    }
}
