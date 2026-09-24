using System.Text.Json.Nodes;

namespace Invigoration.Scr.Aurora;

/// <summary>
/// Answers ConnectionService's echo on SC:R's Aurora WebSocket, whose messages
/// are JSON arrays of <c>[header, body]</c>. The reply's fields (service_id 254,
/// the same token, is_response true, status 0, the received body) are
/// documented. The key names Battle.net uses for the service hash and method on
/// the incoming header are not yet confirmed, so both snake_case and camelCase
/// are accepted until a live capture settles it.
/// </summary>
public static class AuroraKeepalive
{
    public const uint ConnectionServiceHash = 0x65446991;
    public const uint EchoMethod = 3;
    public const uint DisconnectMethod = 4;

    /// <summary>What an incoming Aurora message asks of us.</summary>
    public enum Request
    {
        None,
        Echo,
        Disconnect,
    }

    /// <summary>Classifies an incoming Aurora message. <paramref name="reply"/> is the JSON to send back for an echo, otherwise null.</summary>
    public static Request Handle(string message, out string? reply)
    {
        reply = null;
        if (JsonNode.Parse(message) is not JsonArray { Count: >= 1 } array || array[0] is not JsonObject header)
        {
            return Request.None;
        }

        if (ReadUInt(header, "service_hash", "serviceHash") != ConnectionServiceHash)
        {
            return Request.None;
        }

        switch (ReadUInt(header, "method_id", "methodId"))
        {
            case EchoMethod:
                var response = new JsonObject
                {
                    ["service_id"] = 254,
                    ["token"] = header["token"]?.DeepClone(),
                    ["is_response"] = true,
                    ["status"] = 0,
                };
                reply = new JsonArray(response, array.Count > 1 ? array[1]?.DeepClone() : null).ToJsonString();
                return Request.Echo;
            case DisconnectMethod:
                return Request.Disconnect;
            default:
                return Request.None;
        }
    }

    private static uint? ReadUInt(JsonObject header, params string[] keys)
    {
        foreach (var key in keys)
        {
            if (header[key] is not JsonValue value)
            {
                continue;
            }

            if (value.TryGetValue<uint>(out var number))
            {
                return number;
            }

            if (value.TryGetValue<long>(out var wide) && wide is >= 0 and <= uint.MaxValue)
            {
                return (uint)wide;
            }

            if (value.TryGetValue<string>(out var text) && uint.TryParse(text, out number))
            {
                return number;
            }
        }

        return null;
    }
}
