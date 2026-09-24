using System.Diagnostics;
using System.Text;
using Invigoration.Core.Config;

namespace Invigoration.App.Models;

/// <summary>
/// Runs invigoration-signin, the small macOS app bundled next to Invigoration (see
/// packaging/signin-helper and build-macos.sh), for a Battle.net web sign-in. Avalonia's web
/// view on macOS never reports Battle.net's final redirect to http://localhost:0/?ST=…, so its
/// window spins forever after a correct password; the helper drives WKWebView itself and catches
/// that redirect however WebKit delivers it. Its diagnostics (addresses without queries) go to
/// Logs/sign-in-&lt;time&gt;.log.
/// </summary>
internal static class MacSignInHelper
{
    private const string FileName = "invigoration-signin";
    private const int ExitClosed = 2;

    /// <summary>The helper's path on macOS when it's been bundled; null elsewhere, or in a development run without it.</summary>
    public static string? Find()
    {
        if (!OperatingSystem.IsMacOS())
        {
            return null;
        }

        var path = Path.Combine(AppContext.BaseDirectory, FileName);
        return File.Exists(path) ? path : null;
    }

    /// <summary>
    /// Shows the sign-in and returns the ST credential. Cancelling closes the helper and throws
    /// OperationCanceledException; the user closing its window throws InvalidOperationException,
    /// the same as the built-in window, so callers can tell "never mind" from "declined".
    /// </summary>
    public static async Task<byte[]> RunAsync(string helperPath, Uri challengeUrl, string title, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var start = new ProcessStartInfo(helperPath)
        {
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        start.ArgumentList.Add(challengeUrl.ToString());
        start.ArgumentList.Add("--title");
        start.ArgumentList.Add(title);

        using var trace = OpenTrace();
        trace?.WriteLine($"{DateTime.Now:HH:mm:ss.fff} Sign-in helper starting.");
        using var process = Process.Start(start) ?? throw new InvalidOperationException("The Battle.net sign-in window couldn't be opened.");

        // Its stdin stays open while we wait: the helper quits by itself if Invigoration goes away.
        process.ErrorDataReceived += (_, e) =>
        {
            if (e.Data is { } line)
            {
                lock (process)
                {
                    trace?.WriteLine($"{DateTime.Now:HH:mm:ss.fff} {line}");
                }
            }
        };
        process.BeginErrorReadLine();
        var output = process.StandardOutput.ReadToEndAsync(CancellationToken.None);

        using (cancellationToken.Register(() => Kill(process)))
        {
            await process.WaitForExitAsync(CancellationToken.None).ConfigureAwait(false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        var st = (await output.ConfigureAwait(false))
            .Split('\n')
            .Select(line => line.Trim())
            .FirstOrDefault(line => line.StartsWith("ST=", StringComparison.Ordinal))?[3..];
        lock (process)
        {
            trace?.WriteLine($"{DateTime.Now:HH:mm:ss.fff} Sign-in helper exited ({process.ExitCode}), {(string.IsNullOrEmpty(st) ? "no sign-in" : "signed in")}.");
        }

        if (!string.IsNullOrEmpty(st))
        {
            return Encoding.UTF8.GetBytes(st);
        }

        throw new InvalidOperationException(process.ExitCode == ExitClosed
            ? "The Battle.net login window was closed before finishing."
            : $"The Battle.net sign-in window stopped unexpectedly (exit {process.ExitCode}).");
    }

    private static void Kill(Process process)
    {
        try
        {
            process.Kill();
        }
        catch (InvalidOperationException)
        {
            // Already gone.
        }
    }

    private static StreamWriter? OpenTrace()
    {
        try
        {
            var directory = Path.Combine(ConfigStore.DefaultConfigDirectory(), "Logs");
            Directory.CreateDirectory(directory);
            return new StreamWriter(Path.Combine(directory, $"sign-in-{DateTime.Now:yyyyMMdd-HHmmss}.log"), append: false, Encoding.UTF8) { AutoFlush = true };
        }
        catch (IOException)
        {
            return null;
        }
    }
}
