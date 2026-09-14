using Discord;
using Discord.WebSocket;

namespace Invigoration.Core.Discord;

/// <summary>
/// Thin wrapper around Discord.Net's gateway client for one bot's Discord
/// bridge — logs in, joins the gateway, relays messages from a single
/// configured channel out via <see cref="MessageReceived"/>, and can send
/// text back into that same channel. One instance per <see cref="BotEngine"/>
/// with Config.Discord.Enabled, created/torn down alongside that bot's own
/// connect/disconnect (see BotEngine.Discord.cs) rather than kept running
/// independently.
///
/// Requires the "Message Content" privileged gateway intent to be turned on
/// for the bot application in the Discord Developer Portal — without it,
/// <see cref="SocketMessage.Content"/> comes through empty for messages sent
/// by other users, and nothing will relay. This is a one-time setup step on
/// Discord's side that no amount of code here can turn on remotely.
/// </summary>
public sealed class DiscordBridgeClient : IAsyncDisposable
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(10) };

    private DiscordSocketClient? _client;
    private ulong _channelId;

    // The channel's "Invigoration Relay" webhook, found or created on first use. Once creating it
    // has failed for lack of permission, stay on plain bot messages rather than retrying (and
    // failing) on every single line.
    private string? _webhookUrl;
    private bool _webhookUnavailable;
    private readonly SemaphoreSlim _webhookLock = new(1, 1);

    /// <summary>Fired for a message in the bridged channel from a real (non-bot) user — (username, content).</summary>
    public event Action<string, string>? MessageReceived;

    /// <summary>Diagnostic/error text from the underlying Discord.Net client — wire to LogDebug, not LogInfo, it's chatty.</summary>
    public event Action<string>? Log;

    public async Task StartAsync(string botToken, ulong channelId)
    {
        _channelId = channelId;
        _client = new DiscordSocketClient(new DiscordSocketConfig
        {
            GatewayIntents = GatewayIntents.Guilds | GatewayIntents.GuildMessages | GatewayIntents.MessageContent,
            LogLevel = LogSeverity.Warning,
        });
        _client.Log += OnClientLog;
        _client.MessageReceived += OnMessageReceived;
        await _client.LoginAsync(TokenType.Bot, botToken).ConfigureAwait(false);
        await _client.StartAsync().ConfigureAwait(false);
    }

    private Task OnClientLog(LogMessage msg)
    {
        var exceptionSuffix = msg.Exception is null ? "" : $" ({msg.Exception.Message})";
        Log?.Invoke($"{msg.Severity}: {msg.Message}{exceptionSuffix}");
        return Task.CompletedTask;
    }

    private Task OnMessageReceived(SocketMessage message)
    {
        // Webhook posts are this relay's own Battle.net lines coming back around — never relay them.
        if (message.Author.IsBot || message.Author.IsWebhook || message.Channel.Id != _channelId || string.IsNullOrEmpty(message.Content))
        {
            return Task.CompletedTask;
        }

        MessageReceived?.Invoke(message.Author.Username, message.Content);
        return Task.CompletedTask;
    }

    /// <summary>Posts as the bridge bot itself. Mentions are never honored — Battle.net chat must not be able to ping anyone on Discord.</summary>
    public async Task SendAsync(string text)
    {
        if (_client?.GetChannel(_channelId) is IMessageChannel channel)
        {
            await channel.SendMessageAsync(text, allowedMentions: AllowedMentions.None).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Posts <paramref name="text"/> as <paramref name="username"/>, with <paramref name="avatarUrl"/>
    /// as the avatar, through the channel's relay webhook (see DiscordWebhookIdentity). Falls back
    /// to "**name**: text" from the bot when the webhook can't be used — no Manage Webhooks
    /// permission, a name Discord rejects, or a failed post.
    /// </summary>
    public async Task SendAsAsync(string username, string? avatarUrl, string text)
    {
        if (DiscordWebhookIdentity.IsUsableUsername(username) && await GetWebhookUrlAsync().ConfigureAwait(false) is { } url)
        {
            try
            {
                using var body = new StringContent(DiscordWebhookIdentity.BuildPayload(text, username, avatarUrl), System.Text.Encoding.UTF8, "application/json");
                using var response = await Http.PostAsync(url, body).ConfigureAwait(false);
                if (response.IsSuccessStatusCode)
                {
                    return;
                }

                Log?.Invoke($"Relay webhook post failed ({(int)response.StatusCode}); posting as the bot instead.");
                if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    _webhookUrl = null; // deleted from Discord — find or create it again next time
                }
            }
            catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
            {
                Log?.Invoke($"Relay webhook post failed ({ex.Message}); posting as the bot instead.");
            }
        }

        await SendAsync($"**{username}**: {text}").ConfigureAwait(false);
    }

    private async Task<string?> GetWebhookUrlAsync()
    {
        if (_webhookUrl is not null || _webhookUnavailable)
        {
            return _webhookUrl;
        }

        await _webhookLock.WaitAsync().ConfigureAwait(false);
        try
        {
            if (_webhookUrl is not null || _webhookUnavailable || _client?.GetChannel(_channelId) is not SocketTextChannel channel)
            {
                return _webhookUrl;
            }

            var mine = _client.CurrentUser?.Id;
            var existing = (await channel.GetWebhooksAsync().ConfigureAwait(false))
                .FirstOrDefault(w => w.Name == DiscordWebhookIdentity.WebhookName && w.Token is not null && (w.Creator is null || w.Creator.Id == mine));
            var webhook = existing ?? await channel.CreateWebhookAsync(DiscordWebhookIdentity.WebhookName).ConfigureAwait(false);
            _webhookUrl = $"https://discord.com/api/webhooks/{webhook.Id}/{webhook.Token}";
            return _webhookUrl;
        }
        catch (Exception ex)
        {
            _webhookUnavailable = true;
            Log?.Invoke($"Can't use a relay webhook in this channel ({ex.Message}) — Battle.net names and game logos need the bot to have the Manage Webhooks permission there. Posting as the bot instead.");
            return null;
        }
        finally
        {
            _webhookLock.Release();
        }
    }

    /// <summary>
    /// Sets the bridge bot's own Discord activity/status (e.g. "Playing Invigoration in bnetcc
    /// on useast.battle.net") — the bot's own presence, which Discord.Net's gateway API supports
    /// directly. Deliberately not "set my personal Discord status": automating a real user
    /// account's presence needs a user token and is a Discord self-bot ToS violation regardless
    /// of whose account it is, so that was never on the table here.
    /// </summary>
    public async Task SetPresenceAsync(string activityText)
    {
        if (_client is not null)
        {
            await _client.SetGameAsync(activityText, type: ActivityType.Playing).ConfigureAwait(false);
        }
    }

    public async ValueTask DisposeAsync()
    {
        var client = _client;
        _client = null;
        if (client is null)
        {
            return;
        }

        client.Log -= OnClientLog;
        client.MessageReceived -= OnMessageReceived;
        try
        {
            await client.LogoutAsync().ConfigureAwait(false);
            await client.StopAsync().ConfigureAwait(false);
        }
        catch
        {
            // Best-effort on the way out — the socket may already be dead (e.g. bad token never
            // fully connected), which shouldn't block the rest of this bot's disconnect/dispose.
        }

        client.Dispose();
    }
}
