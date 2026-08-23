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

    private async Task HandleDiscordMessageAsync(string username, string content)
    {
        // Fed through the same pipeline BNCS/Chat-Telnet/SC2 Talk events use — trivia matching
        // and trigger-prefixed command dispatch, in particular — so a Discord user can answer a
        // running trivia round (or run an authorized command) the same as anyone in the actual
        // Battle.net channel. No ChannelIndex: Discord isn't a joined SC2 channel, so this
        // always passes HandleChatEvent's channel-isolation gate, same as a whisper does.
        // Deliberately independent of RelayDiscordToBattlenet below — whether the bot *reacts*
        // to a Discord message and whether that message is *visibly echoed* into Battle.net
        // chat are separate toggles.
        await HandleChatEvent(new ChatEvent(ChatEventType.Talk, username, 0, 0, content, Origin: ChatEventOrigin.Discord)).ConfigureAwait(false);

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
        await SendChatCommandAsync($"[Discord] {username}: {content}").ConfigureAwait(false);
    }

    private async void OnChatMessageForDiscordRelay(ChatEvent chatEvent)
    {
        if (_discordBridge is not { } bridge || !Config.Discord.RelayBattlenetToDiscord ||
            chatEvent.Type is not (ChatEventType.Talk or ChatEventType.Emote))
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
            await bridge.SendAsync($"**{chatEvent.Username}**: {prefix}{chatEvent.Text}{prefix}").ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            LogDebug($"Discord relay send failed: {ex.Message}");
        }
    }
}
