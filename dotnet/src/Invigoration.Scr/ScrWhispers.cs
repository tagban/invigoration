using Invigoration.Sc2.Protobuf;

namespace Invigoration.Scr;

/// <summary>A Battle.net whisper (AuroraChat), addressed to an account rather than a character.</summary>
/// <param name="Outgoing">Battle.net's echo of one we sent (WhisperEchoReceived), rather than one to us.</param>
public sealed record ScrWhisper(uint AccountId, string Text, bool Outgoing);

/// <summary>
/// AuroraChat: Battle.net whispers, which reach a friend whatever game they're in, the Battle.net
/// app or mobile. Method IDs from the retail client's library; the echo as ncarrillo/superiority (MIT)
/// found it.
/// </summary>
public static class ScrWhispers
{
    public const uint Service = 0x924CCFDA;
    public const uint SendWhisperMethod = 0x6251CCD8;
    public const uint WhisperReceivedMethod = 0x7255E575;
    public const uint WhisperEchoReceivedMethod = 0x82B844A8;

    /// <summary>SendWhisper: {1: account ID (fixed32), 2: text}. The client's parser insists on fixed32.</summary>
    public static byte[] SendRequest(uint accountId, string text)
    {
        var request = new ProtoWriter();
        request.WriteFixed32(1, accountId);
        request.WriteString(2, text);
        return request.ToArray();
    }

    /// <summary>WhisperReceived or WhisperEchoReceived: the same shape as the request, with the other account in field 1.</summary>
    public static ScrWhisper? Decode(uint method, byte[] body)
    {
        if (method is not (WhisperReceivedMethod or WhisperEchoReceivedMethod))
        {
            return null;
        }

        uint? accountId = null;
        string? text = null;
        var r = new ProtoReader(body);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1 when type == WireType.Fixed32:
                    accountId = r.ReadFixed32();
                    break;
                case 2 when type == WireType.LengthDelimited:
                    text = r.ReadString();
                    break;
                default:
                    r.Skip(type);
                    break;
            }
        }

        return accountId is { } id && text is not null ? new ScrWhisper(id, text, method == WhisperEchoReceivedMethod) : null;
    }
}
