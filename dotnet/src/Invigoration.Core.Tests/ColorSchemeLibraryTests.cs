using Invigoration.Core.Chat;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

public class ColorSchemeLibraryTests
{
    [Fact]
    public void SaveThenLoad_RoundTripsNameAndColors()
    {
        var dir = ColorSchemeLibrary.Directory;
        var name = $"test-scheme-{Guid.NewGuid():N}";
        var path = Path.Combine(dir, $"{name}.json");
        try
        {
            var scheme = new NamedCustomPalette { Name = name, Colors = ChatPalette.DiabloII.ToCustom() };

            ColorSchemeLibrary.Save(scheme);
            var loaded = ColorSchemeLibrary.Load(path);

            Assert.Equal(name, loaded.Name);
            Assert.Equal(scheme.Colors.Background, loaded.Colors.Background);
            Assert.Equal(scheme.Colors.Highlight, loaded.Colors.Highlight);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ListSchemes_FindsASavedScheme()
    {
        var name = $"test-scheme-{Guid.NewGuid():N}";
        var path = Path.Combine(ColorSchemeLibrary.Directory, $"{name}.json");
        try
        {
            ColorSchemeLibrary.Save(new NamedCustomPalette { Name = name, Colors = new CustomChatPalette() });

            var found = ColorSchemeLibrary.ListSchemes();

            Assert.Contains(found, s => s.Name == name && s.FilePath == path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ToCustom_ThenFromCustom_RoundTripsAllRoles()
    {
        var original = ChatPalette.StarCraft;

        var restored = ChatPalette.FromCustom(original.ToCustom());

        Assert.Equal(original.Background, restored.Background);
        Assert.Equal(original.White, restored.White);
        Assert.Equal(original.Highlight, restored.Highlight);
        Assert.Equal(original.GetUserNameColor(0), restored.GetUserNameColor(0));
    }
}
