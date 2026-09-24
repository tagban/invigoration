using System.Text;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Threading;

namespace Invigoration.App.Models;

/// <summary>
/// Pops Avalonia's native web-auth dialog for a Battle.net login challenge
/// and extracts the resulting "ST" credential — shared between the Battle.net
/// profiles window's explicit "Log in" and every SC2 bot's own challenge handler
/// (set in BotTabViewModel), since both need the exact same popup-then-extract sequence.
/// </summary>
public static class Sc2LoginChallenge
{
    private const string DefaultTitle = "Battle.net Sign-In";

    /// <summary>
    /// Over the app's main window — for a bot, whose own tab may not be showing (auto-connected at
    /// startup, or reconnecting in the background) and so has no view of its own to hang a window on.
    /// The title should say which bot is asking: several can need a sign-in at once, and each saves
    /// whichever account is signed in to its own Battle.net profile.
    /// </summary>
    public static Task<byte[]> ShowAsync(Uri challengeUrl, string title, CancellationToken cancellationToken) =>
        ShowAsync(() => (Application.Current?.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)?.MainWindow, challengeUrl, title, cancellationToken);

    public static Task<byte[]> ShowAsync(TopLevel topLevel, Uri challengeUrl, CancellationToken cancellationToken = default) =>
        ShowAsync(() => topLevel, challengeUrl, DefaultTitle, cancellationToken);

    /// <summary>
    /// Cancelling closes the window (a bot disconnected or removed mid-sign-in) and throws
    /// OperationCanceledException; the user closing it throws InvalidOperationException, so a
    /// caller can tell "never mind" from "declined".
    /// </summary>
    private static async Task<byte[]> ShowAsync(Func<TopLevel?> findOwner, Uri challengeUrl, string title, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        // On macOS, Invigoration's own sign-in app, which catches the final redirect Avalonia's
        // web view never reports. Everywhere else (and in a development run without it), the
        // built-in window below, with the browser-and-paste route beside it.
        if (MacSignInHelper.Find() is { } signInApp)
        {
            return await MacSignInHelper.RunAsync(signInApp, challengeUrl, title, cancellationToken);
        }

        NativeWebDialog? dialog = null;
        BrowserSignInHelper? helper = null;
        var helperClosed = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        // Battle.net finishes by redirecting to http://localhost:0/?ST=…, which the broker is meant to
        // catch. On macOS that redirect arrives as a server-side redirect of the login form's submit,
        // and the web view never reports it as a navigation, so the broker waits forever while the
        // page spins. So the window's address is also watched directly (see SignInTrace), and the
        // credential taken from whichever sees it first.
        var redirected = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
        void Observe(Uri address)
        {
            if (TryReadCredential(address) is { } credential)
            {
                Finish(credential);
            }
        }

        void Finish(string credential)
        {
            if (redirected.TrySetResult(credential))
            {
                Dispatcher.UIThread.Post(CloseDialog);
            }
        }

        var options = new WebAuthenticatorOptions(challengeUrl, new Uri("http://localhost:0/"))
        {
            Mode = WebAuthenticatorMode.NativeWebDialog,

            // Keep Battle.net's web session (its cookies) between sign-ins. Where Battle.net keeps
            // someone signed in ("Keep me logged in"), a later sign-in can then finish without the
            // password. It's the default already; set so a change of default can't quietly lose it.
            NonPersistent = false,

            // The same dialog the broker would make itself, made here only to keep a handle on
            // it: the broker takes no cancellation, so closing the window is the only way to end
            // a sign-in nobody's waiting for any more.
            NativeWebDialogFactory = () =>
            {
                dialog = new NativeWebDialog { Title = title, CanUserResize = true };
                dialog.Resize(600, 700);
                SignInTrace.Attach(dialog, Observe);

                // Beside it, the way round it: the same sign-in in the user's own browser.
                helper = new BrowserSignInHelper(challengeUrl, title, Finish);
                helper.Closed += (_, _) => helperClosed.TrySetResult();
                if (findOwner() is Window owner)
                {
                    helper.Show(owner);
                }
                else
                {
                    helper.Show();
                }

                return dialog;
            },
        };

        // UI thread only (always posted), and once: the registration and the cancel catch below can
        // both get here.
        void CloseDialog()
        {
            var open = dialog;
            dialog = null;
            var openHelper = helper;
            helper = null;
            try
            {
                open?.Close();
            }
            catch (InvalidOperationException)
            {
                // Already closing, or disposed (ObjectDisposedException).
            }

            openHelper?.Close();
        }

        using var closeOnCancel = cancellationToken.Register(() => Dispatcher.UIThread.Post(CloseDialog));

        // FrontClient's own awaits are all ConfigureAwait(false), so by the time this is
        // reached we may already be off the UI thread — and the native dialog's underlying
        // Window can only be created on the UI thread. Hop back for it.
        var authenticating = Dispatcher.UIThread.InvokeAsync(() =>
        {
            // Cancelled while this waited its turn: the close posted then found no window yet.
            if (cancellationToken.IsCancellationRequested)
            {
                return Task.FromCanceled<WebAuthenticationResult>(cancellationToken);
            }

            var owner = findOwner() ?? throw new InvalidOperationException("No window available to show the Battle.net login popup.");
            return WebAuthenticationBroker.AuthenticateAsync(owner, options);
        });

        WebAuthenticationResult result;
        try
        {
            // WaitAsync as well as closing the window: a cancel landing before the window exists
            // can't close it, and the caller shouldn't be held up either way.
            var finished = await Task.WhenAny(authenticating, redirected.Task).WaitAsync(cancellationToken);
            if (finished == redirected.Task)
            {
                // Closing the window ends the broker too; nothing is waiting on how.
                _ = authenticating.ContinueWith(t => _ = t.Exception, TaskScheduler.Default);
                return Encoding.UTF8.GetBytes(await redirected.Task);
            }

            result = await authenticating;
            Dispatcher.UIThread.Post(CloseDialog);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            // Close from here too. Cancellation runs its callbacks newest first, so WaitAsync's own
            // can resume this method, whose `using` then removes closeOnCancel before it has run. The
            // broker makes the window inside the InvokeAsync job, so by the time this post runs, the
            // window exists if it ever will.
            Dispatcher.UIThread.Post(CloseDialog);
            throw new OperationCanceledException(cancellationToken);
        }
        catch (OperationCanceledException)
        {
            // The broker reports the user closing its window as a cancellation of its own. Closing a
            // spinning window is how someone moves on to the browser route, so while that's still
            // open, wait for it rather than giving up.
            if (!helperClosed.Task.IsCompleted
                && await Task.WhenAny(redirected.Task, helperClosed.Task).WaitAsync(cancellationToken) == redirected.Task)
            {
                Dispatcher.UIThread.Post(CloseDialog);
                return Encoding.UTF8.GetBytes(await redirected.Task);
            }

            Dispatcher.UIThread.Post(CloseDialog);
            throw new InvalidOperationException("The Battle.net login window was closed before finishing.");
        }

        if (result.Parameters.TryGetValue("ST", out var st) && !string.IsNullOrEmpty(st))
        {
            return Encoding.UTF8.GetBytes(st);
        }

        throw new InvalidOperationException(result.Error ?? "The Battle.net login window was closed before finishing.");
    }

    /// <summary>The "ST" credential from Battle.net's final redirect to http://localhost:0/, or null for any other address.</summary>
    internal static string? TryReadCredential(Uri address)
    {
        if (!address.IsAbsoluteUri || address.Host != "localhost" || string.IsNullOrEmpty(address.Query))
        {
            return null;
        }

        var st = System.Web.HttpUtility.ParseQueryString(address.Query)["ST"];
        return string.IsNullOrEmpty(st) ? null : st;
    }
}
