namespace Invigoration.Sc2.Front;

/// <summary>
/// The attributes handed over when Front transfers a session to Sunken:
/// where to connect (<see cref="Address"/>) and what to present as proof
/// (<see cref="SessionKey"/>, the 64-byte seed for <c>NativeCrypto</c>'s
/// session-proof derivation), plus the account/game-account identity fields
/// <see cref="ResumeHandshake"/>'s ResumeRequest needs directly.
/// </summary>
public sealed record SunkenHandoff(
    string Address,
    byte[] SessionKey,
    byte AccountRegion,
    string GameAccountName,
    string AccountMail,
    byte[]? LogonResponse)
{
    /// <summary>
    /// Splits <see cref="Address"/> into (host, port), applying
    /// <paramref name="defaultPort"/> when the address carries no port.
    /// Mirrors core/src/bgs/model.rs's NativeHandoff::endpoint, including
    /// bracketed-IPv6 (<c>[::1]:1119</c>) and bare-IPv6 support.
    /// </summary>
    public (string Host, int Port) Endpoint(int defaultPort)
    {
        var value = Address.Trim();
        string host;
        int port;

        if (value.StartsWith('['))
        {
            var closeBracket = value.IndexOf(']');
            if (closeBracket < 0)
            {
                throw new FormatException("SC2 native endpoint has invalid IPv6 syntax.");
            }

            host = value[1..closeBracket];
            var suffix = value[(closeBracket + 1)..];
            port = suffix.Length == 0
                ? defaultPort
                : ParsePort(suffix.StartsWith(':') ? suffix[1..] : throw new FormatException("SC2 native endpoint has invalid IPv6 syntax."));
        }
        else if (value.Count(c => c == ':') == 1)
        {
            var separator = value.LastIndexOf(':');
            host = value[..separator];
            port = ParsePort(value[(separator + 1)..]);
        }
        else
        {
            host = value;
            port = defaultPort;
        }

        if (host.Length == 0 || host.Any(char.IsWhiteSpace))
        {
            throw new FormatException("SC2 native endpoint has an invalid host.");
        }

        if (host.Contains(':') && !System.Net.IPAddress.TryParse(host, out _))
        {
            throw new FormatException("SC2 native endpoint has invalid IPv6 syntax.");
        }

        return (host, port);
    }

    private static int ParsePort(string text) =>
        ushort.TryParse(text, out var port) ? port : throw new FormatException("SC2 native endpoint has an invalid port.");
}

/// <summary>
/// Builds and parses GameUtilities.ProcessClientRequest (Front method 1 on
/// the GameUtilities service) — the call that starts the SC2 Front
/// bootstrap right after Front login and hands off to Sunken. Attribute
/// names, types, and the required/returned sets are exactly as documented
/// at https://superioritybot.com/PROTOCOL's Front RPC section.
/// </summary>
public static class GameUtilitiesSession
{
    /// <summary>
    /// Builds the request. Per upstream, this omits the host-process,
    /// Battle.net-account, program, and client-info fields — only
    /// <see cref="ClientRequest.GameAccountId"/> (the first game-account
    /// entity from Front's LogonResult) and the four attributes are sent.
    /// </summary>
    public static ClientRequest BuildProcessClientRequest(EntityId gameAccountId, byte[] sessionKey, string environment = "US", string locale = "enUS")
    {
        if (sessionKey.Length != 64)
        {
            throw new ArgumentException("Front session key must be 64 bytes.", nameof(sessionKey));
        }

        return new ClientRequest
        {
            GameAccountId = gameAccountId,
            Attributes =
            [
                new Attribute { Name = "LogonTokenRequest", Value = new Variant { StringValue = "0.0.1" } },
                new Attribute { Name = "environment", Value = new Variant { StringValue = environment } },
                new Attribute { Name = "session_key", Value = new Variant { BlobValue = sessionKey } },
                new Attribute { Name = "locale", Value = new Variant { StringValue = locale } },
            ],
        };
    }

    public static SunkenHandoff ParseHandoff(ClientResponse response)
    {
        string? address = null;
        string? gameAccountName = null;
        string? accountMail = null;
        byte[]? sessionKey = null;
        byte[]? logonResponse = null;
        byte? accountRegion = null;

        foreach (var attribute in response.Attributes)
        {
            switch (attribute.Name)
            {
                case "address": address = attribute.Value.StringValue; break;
                case "session_key": sessionKey = attribute.Value.BlobValue; break;
                case "account_region": accountRegion = (byte?)attribute.Value.UintValue; break;
                case "game_account_name": gameAccountName = attribute.Value.StringValue; break;
                case "account_mail": accountMail = attribute.Value.StringValue; break;
                case "logon_response": logonResponse = attribute.Value.BlobValue; break;
            }
        }

        if (address is null || sessionKey is null || accountRegion is null || gameAccountName is null || accountMail is null)
        {
            throw new InvalidOperationException("GameUtilities response is missing a required Sunken handoff attribute.");
        }

        if (address.Length is 0 or > 255)
        {
            throw new InvalidOperationException("SC2 native endpoint has an invalid length.");
        }

        if (sessionKey.Length != 64)
        {
            throw new InvalidOperationException($"Sunken session key must be 64 bytes, was {sessionKey.Length}.");
        }

        if (gameAccountName.Length is 0 or > 32)
        {
            throw new InvalidOperationException("SC2 game-account name has an invalid length.");
        }

        if (accountMail.Length is 0 or > 320)
        {
            throw new InvalidOperationException("SC2 account mail has an invalid length.");
        }

        return new SunkenHandoff(address, sessionKey, accountRegion.Value, gameAccountName, accountMail, logonResponse);
    }
}
