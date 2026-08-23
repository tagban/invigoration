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
    private DiscordSocketClient? _client;
    private ulong _channelId;

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
        if (message.Author.IsBot || message.Channel.Id != _channelId || string.IsNullOrEmpty(message.Content))
        {
            return Task.CompletedTask;
        }

        MessageReceived?.Invoke(message.Author.Username, message.Content);
        return Task.CompletedTask;
    }

    public async Task SendAsync(string text)
    {
        if (_client?.GetChannel(_channelId) is IMessageChannel channel)
        {
            await channel.SendMessageAsync(text).ConfigureAwait(false);
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
