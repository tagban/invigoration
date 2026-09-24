using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace Invigoration.App.Models;

/// <summary>
/// A way round the built-in Battle.net sign-in window when it won't finish (on macOS it can spin
/// forever after a correct password): open the same sign-in in the user's own browser, then paste
/// the address the browser ends up on. That address is Battle.net's http://localhost:0/?ST=…
/// redirect, which no browser can load, so it stays in the address bar to copy.
/// </summary>
internal sealed class BrowserSignInHelper : Window
{
    private readonly Action<string> _onCredential;
    private readonly TextBox _address;
    private readonly TextBlock _problem;

    public BrowserSignInHelper(Uri challengeUrl, string title, Action<string> onCredential)
    {
        _onCredential = onCredential;
        Title = $"{title}: sign in with your browser";
        Width = 520;
        SizeToContent = SizeToContent.Height;
        CanResize = false;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;

        var open = new Button { Content = "Open the sign-in in my browser" };
        open.Click += async (_, _) =>
        {
            if (GetTopLevel(this)?.Launcher is { } launcher)
            {
                await launcher.LaunchUriAsync(challengeUrl);
            }
        };

        _address = new TextBox { PlaceholderText = "http://localhost:0/?ST=…", AcceptsReturn = false };
        _address.TextChanged += (_, _) => TryFinish(quiet: true);
        _problem = new TextBlock { Foreground = Brushes.IndianRed, TextWrapping = TextWrapping.Wrap, IsVisible = false };
        var finish = new Button { Content = "Continue", IsDefault = true, HorizontalAlignment = HorizontalAlignment.Right };
        finish.Click += (_, _) => TryFinish(quiet: false);

        Content = new StackPanel
        {
            Margin = new Thickness(16),
            Spacing = 10,
            Children =
            {
                Text("If the Battle.net window keeps spinning after you log in, sign in with your web browser instead."),
                Text("1. Open the same sign-in in your browser:"),
                open,
                Text("2. Log in. Your browser will then fail to open a page starting with http://localhost:0/. That's expected."),
                Text("3. Copy that whole address from the browser's address bar and paste it here:"),
                _address,
                _problem,
                finish,
            },
        };
    }

    private void TryFinish(bool quiet)
    {
        var text = _address.Text?.Trim() ?? "";
        if (Uri.TryCreate(text, UriKind.Absolute, out var address) && Sc2LoginChallenge.TryReadCredential(address) is { } credential)
        {
            _onCredential(credential);
            return;
        }

        if (!quiet)
        {
            _problem.Text = "That address has no Battle.net sign-in in it. It should start with http://localhost:0/?ST=";
            _problem.IsVisible = true;
        }
    }

    private static TextBlock Text(string text) => new() { Text = text, TextWrapping = TextWrapping.Wrap };
}
