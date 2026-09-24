using System.Text;
using Avalonia.Controls;
using Avalonia.Threading;
using Invigoration.Core.Config;

namespace Invigoration.App.Models;

/// <summary>
/// Writes where the Battle.net sign-in window goes, page by page, to Logs/sign-in-&lt;time&gt;.log in
/// the settings folder: each navigation's start and finish, and any popup the page asks for. Only
/// scheme, host and path are written. Query strings are left out on purpose, since the last
/// redirect carries the sign-in credential in its query.
/// Every address seen, from those events and from checking the window's current address a few
/// times a second, is also handed to <c>observe</c>: that's how the final redirect is caught when the
/// web view doesn't report it as a navigation.
/// </summary>
internal static class SignInTrace
{
    public static void Attach(NativeWebDialog dialog, Action<Uri> observe)
    {
        StreamWriter? log;
        try
        {
            var directory = Path.Combine(ConfigStore.DefaultConfigDirectory(), "Logs");
            Directory.CreateDirectory(directory);
            log = new StreamWriter(Path.Combine(directory, $"sign-in-{DateTime.Now:yyyyMMdd-HHmmss}.log"), append: false, Encoding.UTF8) { AutoFlush = true };
        }
        catch (IOException)
        {
            log = null;
        }

        void Write(string line)
        {
            try
            {
                log?.WriteLine($"{DateTime.Now:HH:mm:ss.fff} {line}");
            }
            catch (ObjectDisposedException)
            {
            }
        }

        void Safely(Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                try
                {
                    log?.WriteLine($"{DateTime.Now:HH:mm:ss.fff} (trace error: {ex.GetType().Name}: {ex.Message})");
                }
                catch
                {
                    // Nothing left to report to.
                }
            }
        }

        void See(Uri? address)
        {
            if (address is not null)
            {
                observe(address);
            }
        }

        Uri? lastAddress = null;
        var poll = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
        poll.Tick += (_, _) => Safely(() =>
        {
            var address = dialog.Source;
            if (address is not null && address != lastAddress)
            {
                lastAddress = address;
                Write($"address is now {Describe(address)}");
                See(address);
            }
        });
        poll.Start();

        // What the page itself says it's doing, every two seconds: its address (no query) and the
        // start of its visible text. Input values aren't part of visible text.
        var peek = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
        string? lastPeek = null;
        peek.Tick += async (_, _) =>
        {
            try
            {
                var state = await dialog.InvokeScript(
                    "JSON.stringify({ at: location.protocol + '//' + location.host + location.pathname, ready: document.readyState, " +
                    "text: (document.body ? document.body.innerText : '').replace(/\\s+/g, ' ').slice(0, 160) })");
                if (state != lastPeek)
                {
                    lastPeek = state;
                    Write($"page: {state}");
                }
            }
            catch (Exception ex)
            {
                Safely(() => Write($"page: (couldn't read: {ex.GetType().Name})"));
            }
        };
        peek.Start();

        Write("Sign-in window opened.");
        // A trace must never take the app down: every handler is fenced off (one that wasn't, on the
        // page's request stream, crashed the app mid sign-in).
        dialog.NavigationStarted += (_, e) => Safely(() =>
        {
            Write($"navigating to {Describe(e.Request)}");
            See(e.Request);
        });
        dialog.NavigationCompleted += (_, e) => Safely(() =>
        {
            Write($"{(e.IsSuccess ? "loaded" : "FAILED to load")} {Describe(e.Request)}");
            See(e.Request);
        });
        dialog.NewWindowRequested += (_, e) => Safely(() =>
        {
            Write($"page asked for a NEW WINDOW: {Describe(e.Request)} (handled: {e.Handled})");
            See(e.Request);
        });
        dialog.Closing += (_, _) =>
        {
            poll.Stop();
            peek.Stop();
            Write("Sign-in window closed.");
            log?.Dispose();
            log = null;
        };
    }

    private static string Describe(Uri? uri) =>
        uri is null ? "(no address)" : $"{uri.Scheme}://{uri.Authority}{uri.AbsolutePath}{(string.IsNullOrEmpty(uri.Query) ? "" : " (with query)")}";
}
