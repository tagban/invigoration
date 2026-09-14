using Invigoration.Core.Chat;
using Invigoration.Core.Discord;

namespace Invigoration.App.Models;

/// <summary>
/// How a Discord user's message looks in a chat log: the Discord logo in the icon slot and the
/// Discord user's own name as the speaker — whether this bot is the one relaying (the message
/// arrived from Discord) or another Invigoration bot relayed it into the channel
/// (<see cref="DiscordRelayLine"/>). With chat icons turned off there's no logo to carry the
/// meaning, so the "[Discord]" text stays on the name instead.
/// </summary>
public static class DiscordRelayRendering
{
    /// <param name="relayedBy">The Battle.net account that actually posted the line, when it's someone else's relay. Shown dimmed after the message: anyone can type "[Discord] name: ..." into a channel, and the logo shouldn't make a spoof look more official than the real sender.</param>
    public static ChatLineViewModel Build(string discordUser, string message, string? relayedBy, ChatPalette palette, bool showIcons)
    {
        var name = showIcons ? discordUser : DiscordRelayLine.SpeakerName(discordUser);
        var segments = new List<ChatLogSegment> { new(palette.GetUserNameColor(0), $"{name}: ") };
        segments.AddRange(ChatColorFormatter.Parse(message, palette.GetChatColor(0), palette));
        if (relayedBy is not null)
        {
            segments.Add(new ChatLogSegment(palette.Gray, $"  (via {relayedBy})"));
        }

        return new ChatLineViewModel(segments, showIcons ? GameIconLoader.Get("discord-relay") : null);
    }
}
