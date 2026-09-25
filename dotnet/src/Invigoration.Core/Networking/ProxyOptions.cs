namespace Invigoration.Core.Networking;

public enum ProxyProtocol
{
    Socks5,
    Http,

    /// <summary>SOCKS4, or SOCKS4a for a host name. No password: a username is sent as the user ID.</summary>
    Socks4,
}

/// <summary>Proxy a bot's connections should tunnel through, so different bots can egress with different apparent source IPs — the only client-side lever against a third-party server's per-IP connection/flood limits.</summary>
public sealed record ProxyOptions(ProxyProtocol Protocol, string Host, int Port, string? Username = null, string? Password = null)
{
    public bool HasCredentials => !string.IsNullOrEmpty(Username);
}
