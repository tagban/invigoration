using System.Text;
using Avalonia.Controls;
using Avalonia.Threading;

namespace Invigoration.App.Models;

/// <summary>
/// Pops Avalonia's native web-auth dialog for a Battle.net login challenge
/// and extracts the resulting "ST" credential — shared between
/// <c>ConfigViewModel</c>'s explicit "Log in" button and <c>BotTabView</c>'s
/// challenge handler for a real SC2 connect, since both need the exact same
/// popup-then-extract sequence.
/// </summary>
public static class Sc2LoginChallenge
{
    public static async Task<byte[]> ShowAsync(TopLevel topLevel, Uri challengeUrl)
    {
        var options = new WebAuthenticatorOptions(challengeUrl, new Uri("http://localhost:0/"))
        {
            Mode = WebAuthenticatorMode.NativeWebDialog,
        };

        // FrontClient's own awaits are all ConfigureAwait(false), so by the time this is
        // reached we may already be off the UI thread — and the native dialog's underlying
        // Window can only be created on the UI thread. Hop back for it.
        var result = await Dispatcher.UIThread.InvokeAsync(
            () => WebAuthenticationBroker.AuthenticateAsync(topLevel, options));

        if (result.Parameters.TryGetValue("ST", out var st) && !string.IsNullOrEmpty(st))
        {
            return Encoding.UTF8.GetBytes(st);
        }

        throw new InvalidOperationException(result.Error ?? "The Battle.net login window was closed before finishing.");
    }
}
