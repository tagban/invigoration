using System.Text;

namespace Invigoration.Sc2.Connection;

/// <summary>
/// Extracts the login credential from Battle.net's web-auth completion
/// signal. After a successful login, the challenge page navigates to
/// <c>http://localhost:0/?ST=&lt;credential&gt;</c> — not a real request (nothing
/// listens on port 0), just a value the embedded browser control is meant
/// to intercept before it tries to actually navigate there. Ported from
/// core/examples/sc2-tui-bot.rs's web_auth_credential, which is the only
/// place in the reference repo that shows this mechanism concretely (it's
/// app-specific, not part of the core protocol crate).
/// </summary>
public static class WebAuthRedirect
{
    public static SecretBytes? TryExtractCredential(string location)
    {
        if (!Uri.TryCreate(location, UriKind.Absolute, out var url))
        {
            return null;
        }

        if (url.Scheme != "http" || url.Host != "localhost" || url.Port != 0)
        {
            return null;
        }

        var query = url.Query.TrimStart('?');
        string? credential = null;
        foreach (var pair in query.Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            var parts = pair.Split('=', 2);
            if (Uri.UnescapeDataString(parts[0]) == "ST" && parts.Length == 2)
            {
                credential = Uri.UnescapeDataString(parts[1]);
                break;
            }
        }

        return credential is null ? null : SecretBytes.TryCreate(Encoding.UTF8.GetBytes(credential));
    }
}
