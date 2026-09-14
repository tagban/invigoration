using Invigoration.Core.Chat;
using Invigoration.Core.Discord;

namespace Invigoration.Core;

/// <summary>
/// Relays chat between this bot's Battle.net channel and a Discord channel,
/// per Config.Discord (see DiscordBridgeConfig's remarks). Started/stopped
/// alongside this bot's own connect/disconnect rather than run independently
/// — there's no separate "connect Discord" action, turning it on in the
/// config is enough. Each direction gets its own flood-protection delay
/// (Config.Discord.RelayDelaySeconds), separate from the general BNCS/Chat
/// SendChatCommandAsync gate, since a burst of Discord messages shouldn't be
/// throttled by (or share a clock with) Battle.net-side chat activity.
/// </summary>
public sealed partial class BotEngine
{
    private DiscordBridgeClient? _discordBridge;
    private DateTime _nextDiscordToBattlenetAllowedUtc = DateTime.MinValue;
    private DateTime _nextBattlenetToDiscordAllowedUtc = DateTime.MinValue;

    private void WireDiscordBridge() => ChatMessage += OnChatMessageForDiscordRelay;

    /// <summary>
    /// Fire-and-forget on purpose: connecting to Discord's gateway is a
    /// separate, independently-slow network operation from the Battle.net
    /// connect this runs alongside (see ConnectAsync) — it shouldn't delay
    /// or fail that connection if Discord is slow, unreachable, or the token
    /// is bad.
    /// </summary>
    private void StartDiscordBridgeIfEnabled()
    {
        if (!Config.Discord.Enabled || string.IsNullOrWhiteSpace(Config.Discord.BotToken))
        {
            return;
        }

        SafeFireAndForget(ConnectDiscordBridgeAsync(), "connecting the Discord bridge");
    }

    private async Task ConnectDiscordBridgeAsync()
    {
        var bridge = new DiscordBridgeClient();
        bridge.Log += msg => LogDebug($"Discord: {msg}");
        bridge.MessageReceived += (username, content) =>
            SafeFireAndForget(HandleDiscordMessageAsync(username, content), "relaying a Discord message to Battle.net");

        await bridge.StartAsync(Config.Discord.BotToken, Config.Discord.ChannelId).ConfigureAwait(false);
        _discordBridge = bridge;
        LogInfo("Discord bridge connected.");
        await UpdateDiscordPresenceAsync().ConfigureAwait(false);
    }

    /// <summary>
    /// Refreshes the bridge bot's Discord activity to reflect where this bot actually is right
    /// now — called once the bridge connects, and again from BotEngine.Bncs.cs whenever a
    /// ChatEventType.Channel event updates _session.CurrentChannelName (join/rejoin/channel
    /// change), so the presence text doesn't go stale. A no-op if the bridge isn't connected.
    /// </summary>
    private async Task UpdateDiscordPresenceAsync()
    {
        if (_discordBridge is not { } bridge)
        {
            return;
        }

        var activityText = string.IsNullOrEmpty(_session.CurrentChannelName)
            ? $"Invigoration on {Config.BattlenetServer}"
            : $"Invigoration in {_session.CurrentChannelName} on {Config.BattlenetServer}";
        try
        {
            await bridge.SetPresenceAsync(activityText).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            LogDebug($"Discord presence update failed: {ex.Message}");
        }
    }

    private async Task StopDiscordBridgeAsync()
    {
        var bridge = _discordBridge;
        _discordBridge = null;
        if (bridge is not null)
        {
            await bridge.DisposeAsync().ConfigureAwait(false);
        }
    }

    /// <summary>
    /// How a Discord user is named everywhere on this side of the bridge: "[Discord] name". It's
    /// both the label people see and the identity the bot reasons with — and because Battle.net
    /// account names can't contain spaces, it can never collide with a real account. That matters:
    /// Discord names are self-chosen, so under their bare name someone could register a Discord
    /// account matching the bot master's Battle.net name and run master commands through the
    /// relay, or score trivia points and roster rank as someone else.
    /// </summary>
    public static string DiscordSpeakerName(string discordUsername) => DiscordRelayLine.SpeakerName(discordUsername);

    /// <summary>Only real Battle.net chat goes out to Discord — never a message that came in from Discord in the first place, which would post the user's own message back at them.</summary>
    internal static bool ShouldRelayToDiscord(ChatEvent chatEvent, bool relayBattlenetToDiscord) =>
        relayBattlenetToDiscord &&
        chatEvent.Origin != ChatEventOrigin.Discord &&
        chatEvent.Type is ChatEventType.Talk or ChatEventType.Emote;

    private async Task HandleDiscordMessageAsync(string username, string content)
    {
        var speaker = DiscordSpeakerName(username);

        // Fed through the same pipeline BNCS/Chat-Telnet/SC2 Talk events use — trivia matching
        // and trigger-prefixed command dispatch, in particular — so a Discord user can answer a
        // running trivia round, or use the commands open to everyone (help, trivia score), the
        // same as anyone in the actual Battle.net channel. Commands that need the bot master or a
        // rank never authorize for a "[Discord] name" speaker (see DiscordSpeakerName). No ChannelIndex: Discord isn't a joined SC2 channel, so this
        // always passes HandleChatEvent's channel-isolation gate, same as a whisper does.
        // Deliberately independent of RelayDiscordToBattlenet below — whether the bot *reacts*
        // to a Discord message and whether that message is *visibly echoed* into Battle.net
        // chat are separate toggles.
        await HandleChatEvent(new ChatEvent(ChatEventType.Talk, speaker, 0, 0, content, Origin: ChatEventOrigin.Discord)).ConfigureAwait(false);

        if (!Config.Discord.RelayDiscordToBattlenet)
        {
            return;
        }

        var waitMs = (_nextDiscordToBattlenetAllowedUtc - DateTime.UtcNow).TotalMilliseconds;
        if (waitMs > 0)
        {
            await Task.Delay((int)waitMs).ConfigureAwait(false);
        }

        _nextDiscordToBattlenetAllowedUtc = DateTime.UtcNow.AddSeconds(Math.Max(0, Config.Discord.RelayDelaySeconds));

        // Split here rather than leaving it to SendChatCommandAsync, so every piece of a long
        // message keeps the "[Discord] name: " lead — that's what lets other Invigoration bots in
        // the channel show each piece under the Discord user's name (DiscordRelayLine.TryParse).
        // No local echo: the Discord message already shows in this bot's chat log once, from the
        // HandleChatEvent above; echoing the relay too used to print it a second time as
        // "<bot>: [Discord] name: message".
        var lead = DiscordRelayLine.Format(username, "");
        foreach (var piece in Chat.ChatLineSplitter.Split(content, Chat.ChatLineSplitter.MaxLineLength - lead.Length))
        {
            await SendChatCommandAsync(lead + piece, sc2ChannelOverride: null, echoLocally: false).ConfigureAwait(false);
        }
    }

    private async void OnChatMessageForDiscordRelay(ChatEvent chatEvent)
    {
        if (_discordBridge is not { } bridge || !ShouldRelayToDiscord(chatEvent, Config.Discord.RelayBattlenetToDiscord))
        {
            return;
        }

        try
        {
            var waitMs = (_nextBattlenetToDiscordAllowedUtc - DateTime.UtcNow).TotalMilliseconds;
            if (waitMs > 0)
            {
                await Task.Delay((int)waitMs).ConfigureAwait(false);
            }

            _nextBattlenetToDiscordAllowedUtc = DateTime.UtcNow.AddSeconds(Math.Max(0, Config.Discord.RelayDelaySeconds));
            var prefix = chatEvent.Type == ChatEventType.Emote ? "*" : "";
            var text = $"{prefix}{chatEvent.Text}{prefix}";
            if (Config.Discord.PostAsBattlenetUsers)
            {
                // Posted as the speaker, with their game's classic logo as the avatar — see DiscordWebhookIdentity.
                _lastKnownProduct.TryGetValue(chatEvent.Username, out var product);
                await bridge.SendAsAsync(chatEvent.Username, DiscordWebhookIdentity.AvatarUrlFor(product), text).ConfigureAwait(false);
            }
            else
            {
                await bridge.SendAsync($"**{chatEvent.Username}**: {text}").ConfigureAwait(false);
            }
        }
        catch (Exception ex)
        {
            LogDebug($"Discord relay send failed: {ex.Message}");
        }
    }
}
