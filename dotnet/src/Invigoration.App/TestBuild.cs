using Invigoration.Core.Config;

namespace Invigoration.App;

/// <summary>What makes the test build different: its own settings folder.</summary>
public static class TestBuild
{
    /// <summary>Names the settings folder (under Application Support) the test build uses instead of "Invigoration".</summary>
    public const string SettingsFolderVariable = "INVIGORATION_SETTINGS_FOLDER";

    public static bool IsActive => ConfigStore.DirectoryOverride is not null;

    /// <summary>Added to the window title so the test copy is never mistaken for the everyday one.</summary>
    public static string TitleSuffix => IsActive ? " (Test build)" : "";
}
