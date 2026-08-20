using Invigoration.Sc2.Protobuf;

namespace Invigoration.Sc2.Front;

/// <summary>bgs.protocol.challenge.v1 messages (CAPTCHA/external challenge notifications during logon).</summary>
public sealed class ChallengeExternalRequest
{
    public string? RequestToken { get; init; }
    public string? PayloadType { get; init; }
    public byte[]? Payload { get; init; }

    public static ChallengeExternalRequest Decode(byte[] data)
    {
        string? requestToken = null, payloadType = null;
        byte[]? payload = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: requestToken = r.ReadString(); break;
                case 2: payloadType = r.ReadString(); break;
                case 3: payload = r.ReadLengthDelimited(); break;
                default: r.Skip(type); break;
            }
        }

        return new ChallengeExternalRequest { RequestToken = requestToken, PayloadType = payloadType, Payload = payload };
    }

    /// <summary>
    /// Extracts and validates the web-authentication URL this challenge
    /// carries, before it's ever shown to a WebView/WebAuthenticationBroker.
    /// Mirrors core/src/bgs/model.rs's challenge_url: only an https URL on
    /// an ".account.battle.net" host is accepted — this is what stops a
    /// compromised or spoofed Front connection from directing the login
    /// flow at an arbitrary phishing page.
    /// </summary>
    public Uri GetValidatedWebAuthUrl()
    {
        if (PayloadType != "web_auth_url")
        {
            throw new InvalidOperationException($"Unsupported external challenge payload type '{PayloadType}'.");
        }

        if (Payload is null)
        {
            throw new InvalidOperationException("External challenge has no payload.");
        }

        Uri url;
        try
        {
            url = new Uri(System.Text.Encoding.UTF8.GetString(Payload));
        }
        catch (Exception ex) when (ex is FormatException or ArgumentException)
        {
            throw new InvalidOperationException("Web authentication URL is not valid UTF-8/URI.", ex);
        }

        if (url.Scheme != "https" || !url.Host.EndsWith(".account.battle.net", StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Battle.net returned an unexpected authentication URL.");
        }

        return url;
    }
}

public sealed class ChallengeExternalResult
{
    public string? RequestToken { get; init; }
    public bool? Passed { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteString(1, RequestToken);
        w.WriteBool(2, Passed);
        return w.ToArray();
    }

    public static ChallengeExternalResult Decode(byte[] data)
    {
        string? requestToken = null;
        bool? passed = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: requestToken = r.ReadString(); break;
                case 2: passed = r.ReadVarint() != 0; break;
                default: r.Skip(type); break;
            }
        }

        return new ChallengeExternalResult { RequestToken = requestToken, Passed = passed };
    }
}
