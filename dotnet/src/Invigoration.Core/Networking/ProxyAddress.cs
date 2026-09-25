using System.Globalization;
using System.Net;

namespace Invigoration.Core.Networking;

/// <summary>A proxy address as pasted: host and port, maybe a login, and the protocol if the text named one.</summary>
public sealed record ProxyAddress(ProxyProtocol? Protocol, string Host, int Port, string? Username, string? Password);

/// <summary>
/// Reads a proxy address in the forms proxy providers hand out:
/// <list type="bullet">
/// <item>a URL: <c>socks5://user:pass@host:port</c>, <c>socks5h://</c>, <c>socks4://</c>, <c>socks4a://</c>, <c>http://</c> (user and password may be %-encoded)</item>
/// <item><c>user:pass@host:port</c></item>
/// <item><c>user:pass:host:port</c> or <c>host:port:user:pass</c>, told apart by which part is the port</item>
/// <item><c>host:port</c>, with an IPv6 host in brackets</item>
/// </list>
/// </summary>
public static class ProxyAddressParser
{
    public static bool TryParse(string? text, out ProxyAddress address, out string error)
    {
        address = null!;
        error = "";
        var value = text?.Trim() ?? "";
        if (value.Length == 0)
        {
            error = "Paste a proxy address.";
            return false;
        }

        ProxyProtocol? protocol = null;
        var scheme = value.IndexOf("://", StringComparison.Ordinal);
        if (scheme >= 0)
        {
            var name = value[..scheme].ToLowerInvariant();
            protocol = name switch
            {
                "socks5" or "socks5h" or "socks" => ProxyProtocol.Socks5,
                "socks4" or "socks4a" => ProxyProtocol.Socks4,
                "http" => ProxyProtocol.Http,
                _ => null,
            };
            if (protocol is null)
            {
                error = name == "https"
                    ? "https:// proxies (TLS to the proxy itself) aren't supported; most providers' HTTP proxies work as http://."
                    : $"Unknown proxy type \"{name}\". Use socks5://, socks4:// or http://.";
                return false;
            }

            value = value[(scheme + 3)..].TrimEnd('/');
        }

        string? user = null, pass = null;
        string hostPort;
        var at = value.LastIndexOf('@');
        if (at >= 0)
        {
            (user, pass) = SplitLogin(value[..at]);
            hostPort = value[(at + 1)..];
        }
        else if (!value.StartsWith('[') && value.Count(c => c == ':') >= 3)
        {
            // Four colon-separated parts: which pair is host:port?
            var parts = value.Split(':');
            if (parts.Any(p => p.Length == 0))
            {
                error = "An IPv6 proxy address needs brackets, like [2001:db8::1]:1080.";
                return false;
            }

            var hostFirst = IsPort(parts[1]) && (!IsPort(parts[^1]) || LooksLikeHost(parts[0]) || !LooksLikeHost(parts[^2]));
            if (hostFirst && IsHostName(parts[0]))
            {
                hostPort = $"{parts[0]}:{parts[1]}";
                user = parts[2];
                pass = string.Join(":", parts[3..]);
            }
            else if (IsPort(parts[^1]) && IsHostName(parts[^2]))
            {
                hostPort = $"{parts[^2]}:{parts[^1]}";
                user = parts[0];
                pass = string.Join(":", parts[1..^2]);
            }
            else
            {
                error = "Couldn't tell which part is the host and port. Try user:pass@host:port.";
                return false;
            }
        }
        else
        {
            hostPort = value;
        }

        if (!TrySplitHostPort(hostPort, out var host, out var port))
        {
            error = "Expected host:port, e.g. 203.0.113.5:1080.";
            return false;
        }

        address = new ProxyAddress(protocol, host, port, Empty(user), Empty(pass));
        return true;
    }

    /// <summary>A one-line description of a parsed address, without the password.</summary>
    public static string Describe(ProxyAddress address, ProxyProtocol fallback) =>
        $"{Name(address.Protocol ?? fallback)} {address.Host}:{address.Port}" +
        (address.Username is { } user ? $", signing in as {user}{(address.Password is null ? "" : " with a password")}" : "");

    public static string Name(ProxyProtocol protocol) => protocol switch
    {
        ProxyProtocol.Socks5 => "SOCKS5",
        ProxyProtocol.Socks4 => "SOCKS4",
        _ => "HTTP",
    };

    private static (string? User, string? Pass) SplitLogin(string login)
    {
        var colon = login.IndexOf(':');
        return colon < 0
            ? (Decode(login), null)
            : (Decode(login[..colon]), Decode(login[(colon + 1)..]));
    }

    private static bool TrySplitHostPort(string text, out string host, out int port)
    {
        host = "";
        port = 0;
        string portText;
        if (text.StartsWith('['))
        {
            var close = text.IndexOf(']');
            if (close < 0 || close + 1 >= text.Length || text[close + 1] != ':')
            {
                return false;
            }

            host = text[1..close];
            portText = text[(close + 2)..];
        }
        else
        {
            var colon = text.LastIndexOf(':');
            if (colon <= 0)
            {
                return false;
            }

            host = text[..colon];
            portText = text[(colon + 1)..];
        }

        // An unbracketed host can't hold a colon: that would be an IPv6 address without its brackets.
        if (host.Length == 0 || (!text.StartsWith('[') && host.Contains(':')) || !IsPort(portText))
        {
            return false;
        }

        port = int.Parse(portText, CultureInfo.InvariantCulture);
        return true;
    }

    private static bool IsPort(string text) =>
        int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var port) && port is > 0 and <= 65535;

    private static bool LooksLikeHost(string text) =>
        IPAddress.TryParse(text, out _) || (text.Contains('.') && !text.StartsWith('.') && !text.EndsWith('.'));

    /// <summary>Could be a host: an address, or a name with a letter in it (not a bare number).</summary>
    private static bool IsHostName(string text) => LooksLikeHost(text) || text.Any(char.IsLetter);

    private static string Decode(string text) => Uri.UnescapeDataString(text);

    private static string? Empty(string? text) => string.IsNullOrEmpty(text) ? null : text;
}
