using Invigoration.Core.Music;

namespace Invigoration.Core;

/// <summary>
/// Auto-away messaging: after Config.IdleMinutes of no chat activity (heard or sent) in this
/// bot's channel, sends Config.IdleMessage once — doesn't repeat until real activity resumes and
/// a new idle period elapses. Off by default (IdleMinutes 0, or an empty IdleMessage). Distinct
/// from KeepAliveInterval's SID_NULL heartbeat (BotEngine.cs) — that's a protocol-level "don't
/// let the connection get dropped" ping, this is a visible, chat-level "I've gone quiet" line.
/// Classic BNCS/Chat-Telnet only for now (wired alongside StartKeepAlive/StopKeepAlive in
/// OnBncsConnected/_bncs.Disconnected) — SC2/SC:R/WC3:R's multi-channel model doesn't map onto
/// "the channel" the same way a single classic-BNCS channel does.
/// </summary>
public sealed partial class BotEngine
{
    private static readonly TimeSpan IdleCheckInterval = TimeSpan.FromSeconds(30);

    private DateTimeOffset _lastChatActivityUtc = DateTimeOffset.UtcNow;
    private bool _idleMessageSent;
    private CancellationTokenSource? _idleCheckCts;

    /// <summary>Resets the idle clock — called on any real chat activity, heard (HandleChatEvent's Talk/Emote/Whisper cases) or sent (SendChatCommandAsync) — and clears the "already sent" flag so the next idle period can trigger again.</summary>
    private void NoteChatActivity()
    {
        _lastChatActivityUtc = DateTimeOffset.UtcNow;
        _idleMessageSent = false;
    }

    private void StartIdleWatcher()
    {
        StopIdleWatcher();
        _idleCheckCts = new CancellationTokenSource();
        SafeFireAndForget(RunIdleWatcherLoopAsync(_idleCheckCts.Token), "checking idle state");
    }

    private void StopIdleWatcher()
    {
        _idleCheckCts?.Cancel();
        _idleCheckCts?.Dispose();
        _idleCheckCts = null;
    }

    private async Task RunIdleWatcherLoopAsync(CancellationToken cancellationToken)
    {
        try
        {
            using var timer = new PeriodicTimer(IdleCheckInterval);
            while (await timer.WaitForNextTickAsync(cancellationToken).ConfigureAwait(false))
            {
                await CheckIdleAsync().ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException)
        {
            // Normal on disconnect — StopIdleWatcher() cancels this loop deliberately.
        }
    }

    private async Task CheckIdleAsync()
    {
        if (_idleMessageSent || Config.IdleMinutes <= 0 || string.IsNullOrEmpty(Config.IdleMessage))
        {
            return;
        }

        if (DateTimeOffset.UtcNow - _lastChatActivityUtc < TimeSpan.FromMinutes(Config.IdleMinutes))
        {
            return;
        }

        // Set before the send (not after): SendChatCommandAsync itself calls NoteChatActivity(),
        // which would otherwise immediately clear this same flag the send is trying to set.
        _idleMessageSent = true;
        await SendChatCommandAsync(await ResolveIdlePlaceholdersAsync(Config.IdleMessage).ConfigureAwait(false)).ConfigureAwait(false);
    }

    /// <summary>
    /// %Ver%/%Uptime%/%MusicPlaying%/%Username% — resolved at send time (not when the message was
    /// typed into the config) so each reflects current state rather than whatever was true when
    /// the idle message was configured. %MusicPlaying% becomes "" (the placeholder just vanishes)
    /// when nothing's playing or the music tab isn't open, rather than some broken-looking error
    /// placeholder text.
    /// </summary>
    private async Task<string> ResolveIdlePlaceholdersAsync(string template)
    {
        NowPlayingInfo? nowPlaying = null;
        if (template.Contains("%MusicPlaying%", StringComparison.OrdinalIgnoreCase) && MusicPlayerRegistry.Controller is { } controller)
        {
            nowPlaying = await controller.GetNowPlayingAsync().ConfigureAwait(false);
        }

        var musicText = nowPlaying is null ? "" : $"{nowPlaying.Title} by {nowPlaying.Artist}";

        return template
            .Replace("%Ver%", AppVersion.Current, StringComparison.OrdinalIgnoreCase)
            .Replace("%Uptime%", FormatUptime(), StringComparison.OrdinalIgnoreCase)
            .Replace("%MusicPlaying%", musicText, StringComparison.OrdinalIgnoreCase)
            .Replace("%Username%", Config.Username, StringComparison.OrdinalIgnoreCase);
    }
}
