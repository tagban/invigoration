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

        NativeWebDialog? dialog = null;
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
                return dialog;
            },
        };

        // UI thread only (always posted), and once: the registration and the cancel catch below can
        // both get here.
        void CloseDialog()
        {
            var open = dialog;
            dialog = null;
            try
            {
                open?.Close();
            }
            catch (InvalidOperationException)
            {
                // Already closing, or disposed (ObjectDisposedException).
            }
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
            result = await authenticating.WaitAsync(cancellationToken);
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
            // The broker reports the user closing its window as a cancellation of its own.
            throw new InvalidOperationException("The Battle.net login window was closed before finishing.");
        }

        if (result.Parameters.TryGetValue("ST", out var st) && !string.IsNullOrEmpty(st))
        {
            return Encoding.UTF8.GetBytes(st);
        }

        throw new InvalidOperationException(result.Error ?? "The Battle.net login window was closed before finishing.");
    }
}
