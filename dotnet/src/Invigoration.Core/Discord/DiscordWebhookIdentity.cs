using System.Text.Json;
using System.Text.Json.Nodes;

namespace Invigoration.Core.Discord;

/// <summary>
/// How a Battle.net user appears on Discord when the relay posts as them through a webhook: their
/// Battle.net name as the message's author, and their game's classic Battle.net logo as its
/// avatar. A bot's own messages always show the bot's name and picture; a webhook message can
/// carry a different name and avatar each time, which is what makes this possible.
/// </summary>
public static class DiscordWebhookIdentity
{
    /// <summary>
    /// Discord only takes a per-message avatar as a URL, so the logos are served from this repo:
    /// square 128×128 copies of the classic 28×14 icons (docs/discord-avatars), scaled up crisply
    /// and centered so Discord's round crop doesn't clip them.
    /// </summary>
    public const string AvatarBaseUrl = "https://raw.githubusercontent.com/tagban/invigoration/main/docs/discord-avatars/";

    public const string WebhookName = "Invigoration Relay";

    private const int MaxWebhookUsernameLength = 80;

    private static readonly Dictionary<string, string> AvatarFiles = new(StringComparer.Ordinal)
    {
        ["RATS"] = "sc",
        ["PXES"] = "scbw",
        ["RTSJ"] = "jsc",
        ["RHSS"] = "sware",
        ["NB2W"] = "war2",
        ["LTRD"] = "diablo",
        ["RHSD"] = "dshr",
        ["VD2D"] = "diablo2",
        ["PX2D"] = "d2exp",
        ["3RAW"] = "war3",
        ["PX3W"] = "w3tft",
        ["TAHC"] = "chat",
    };

    /// <summary>The logo URL for a 4-character wire product code (e.g. "RATS"), or null for an unknown or missing product — Discord then shows the webhook's default avatar.</summary>
    public static string? AvatarUrlFor(string? product) =>
        product is { Length: >= 4 } && AvatarFiles.TryGetValue(product[..4], out var file) ? $"{AvatarBaseUrl}{file}.png" : null;

    /// <summary>
    /// Whether Discord will accept this as a webhook message's author name: 1-80 characters, and
    /// never containing "discord" or "clyde" (Discord rejects both). Anything that fails posts in
    /// the plain "**name**: message" form instead.
    /// </summary>
    public static bool IsUsableUsername(string name)
    {
        var trimmed = name.Trim();
        return trimmed.Length is > 0 and <= MaxWebhookUsernameLength &&
               !trimmed.Contains("discord", StringComparison.OrdinalIgnoreCase) &&
               !trimmed.Contains("clyde", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The JSON body for executing the webhook. <c>allowed_mentions</c> is always empty: chat from
    /// Battle.net must never be able to ping @everyone, a role, or a user on the Discord side just
    /// by typing it.
    /// </summary>
    public static string BuildPayload(string content, string username, string? avatarUrl)
    {
        var payload = new JsonObject
        {
            ["content"] = content,
            ["username"] = username.Trim(),
            ["allowed_mentions"] = new JsonObject { ["parse"] = new JsonArray() },
        };
        if (avatarUrl is not null)
        {
            payload["avatar_url"] = avatarUrl;
        }

        return payload.ToJsonString(new JsonSerializerOptions { WriteIndented = false });
    }
}
