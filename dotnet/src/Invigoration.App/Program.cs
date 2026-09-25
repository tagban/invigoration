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
        AppDomain.CurrentDomain.UnhandledException += (_, e) => WriteCrashLog(e.ExceptionObject as Exception);
        try
        {
            BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
        }
        catch (Exception ex)
        {
            // An exception on the UI thread ends up here. The system crash report can't name it.
            WriteCrashLog(ex);
            throw;
        }
    }

    /// <summary>Writes an unhandled exception, with its whole stack, to Logs/crash-&lt;time&gt;.log in the settings folder.</summary>
    private static void WriteCrashLog(Exception? ex)
    {
        if (ex is null)
        {
            return;
        }

        try
        {
            var folder = System.IO.Path.Combine(Invigoration.Core.Config.ConfigStore.DefaultConfigDirectory(), "Logs");
            System.IO.Directory.CreateDirectory(folder);
            System.IO.File.WriteAllText(
                System.IO.Path.Combine(folder, $"crash-{DateTime.Now:yyyyMMdd-HHmmss}.log"),
                $"Invigoration {Invigoration.Core.AppVersion.Current} crashed at {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}{ex}");
        }
        catch (Exception)
        {
            // Nothing more to do while crashing.
        }
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
