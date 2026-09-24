using Avalonia;
using System;

namespace Invigoration.App;

sealed class Program
{
    // Initialization code. Don't use any Avalonia, third-party APIs or any
    // SynchronizationContext-reliant code before AppMain is called: things aren't initialized
    // yet and stuff might break.
    [STAThread]
    public static void Main(string[] args)
    {
        UseTestSettingsFolderIfAsked();
        BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    /// <summary>
    /// The test build ("Invigoration Test.app", see build-macos.sh --test) names its own settings
    /// folder, so it never reads or changes the everyday copy's bots, profiles or sign-ins. Set
    /// before anything loads, since every store reads ConfigStore.DefaultConfigDirectory().
    /// </summary>
    private static void UseTestSettingsFolderIfAsked()
    {
        var name = Environment.GetEnvironmentVariable(TestBuild.SettingsFolderVariable);
        if (string.IsNullOrWhiteSpace(name) || name.IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) >= 0)
        {
            return;
        }

        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        Invigoration.Core.Config.ConfigStore.DirectoryOverride = System.IO.Path.Combine(appData, name);
    }

    // Avalonia configuration, don't remove; also used by visual designer.
    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
#if DEBUG
            .WithDeveloperTools()
#endif
            .WithInterFont()
            .LogToTrace();
}
