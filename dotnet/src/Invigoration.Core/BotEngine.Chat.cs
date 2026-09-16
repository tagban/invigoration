using Invigoration.Core.Chat;
using Invigoration.Core.Networking;

namespace Invigoration.Core;

/// <summary>
/// Battle.net/PVPGN's older plain-text, line-based "Chat" connection type —
/// picked in the Game list as "Chat / Telnet" (<see cref="Protocol.BncsProduct.Chat"/>)
/// rather than as a game, since it isn't tied to one. No BNLS, CD-key, or
/// version-check is involved: login is just a two-prompt (username, then
/// password) exchange — see _chatTelnetPromptsSeen's remarks for why prompts
/// are recognized by shape rather than exact wording — after which every line
/// is a numbered event (see <see cref="ChatTelnetEventParser"/>) fed into the
/// exact same <see cref="HandleChatEvent"/> pipeline the binary protocol uses,
/// so the whole roster/rank/trivia/command-dispatch stack downstream of that
/// method doesn't need to know which transport it came from.
///
/// Because there's no handshake to speak of, this is also the fastest way back
/// onto a private server after a drop — about one round trip, against the
/// several a BNCS logon needs even with cached checks (see BotEngine.Reconnect.cs).
/// Confirmed against a live capture for the login handshake and the
/// USER/JOIN/TALK/CHANNEL events; the rest of the event-type mapping is a
/// best-effort extrapolation of the confirmed pattern (see
/// ChatTelnetEventParser's remarks) — expect to need live-testing fixes.
/// </summary>
public sealed partial class BotEngine
{
    private readonly ChatTelnetConnection _chatTelnet = new();

    /// <summary>How long to wait for a login prompt before deciding this server doesn't send one. Long enough to beat any real banner (they arrive within a round trip), short enough not to blunt a rapid reconnect the first time.</summary>
    private static readonly TimeSpan SilentServerWindow = TimeSpan.FromSeconds(1);

    /// <summary>Set once this server has proved it sends no prompt at all, so later reconnects send the login immediately instead of waiting <see cref="SilentServerWindow"/> out again. Kept for the life of the engine — a server doesn't change its mind mid-session.</summary>
    private bool _chatTelnetServerIsSilent;

    /// <summary>Whether this bot connects over the plain-text Chat/telnet protocol rather than binary BNCS.</summary>
    private bool UsesChatTelnet => Protocol.BncsProduct.IsChatTelnet(Config.Product);

    /// <summary>
    /// How many colon-terminated prompt lines have been seen so far this
    /// login attempt — 0 before the username prompt, 1 after it's answered
    /// (waiting for the password prompt), 2 once both are sent. Counting
    /// prompts by shape (ends with ':') rather than matching specific
    /// wording ("Username:"/"Login:"/"Account name:"/etc.) is deliberate:
    /// a live capture against atlas.bnetdocs.org showed a server whose
    /// banner reads "Enter your login name and password." (not "account
    /// name" like the original sample this was built from) — matching
    /// exact prompt text is fragile across different PVPGN configs, but
    /// every variant seen so far ends an interactive prompt line with ':'
    /// while banner/info sentences end with '.' or a bracketed IP.
    ///
    /// atlas.bnetdocs.org turned out not to send field-specific prompts at
    /// all, though — just that one instructional sentence, then it silently
    /// expects username then password with nothing further in between. See
    /// <see cref="_chatTelnetCredentialsSent"/> for that fallback path.
    /// </summary>
    private int _chatTelnetPromptsSeen;

    /// <summary>
    /// True once username+password have been sent, however that got
    /// triggered — guards against the colon-prompt path and the
    /// instructional-sentence fallback both firing for the same login.
    /// </summary>
    private bool _chatTelnetCredentialsSent;

    private void WireChatTelnet()
    {
        _chatTelnet.Connected += OnChatTelnetConnected;
        _chatTelnet.PacketReceived += frame =>
            SafeFireAndForget(HandleChatTelnetLineAsync(ChatTelnetConnection.DecodeLine(frame)), "handling a Chat-protocol line");
        _chatTelnet.Disconnected += ex =>
        {
            StopIdleWatcher();
            _friends.Clear();
            _auth.LoggedOnToBncs = false;
            _session.CurrentChannelName = "";

            // Same as the BNCS path: an attempt this engine closed itself to open its replacement
            // (rapid reconnect abandoning a half-open attempt) isn't a drop to report or react to.
            if (Interlocked.Exchange(ref _replacingBncsConnection, 0) == 1)
            {
                LogDebug("Replaced the previous Chat connection with a new one.");
                BncsDisconnected?.Invoke(ex);
                return;
            }

            LogError($"Battle.net disconnected{(ex is null ? "." : $": {ex.Message}")}");
            BncsDisconnected?.Invoke(ex);
            MaybeScheduleAutoReconnect();
        };
    }

    private async Task ConnectChatTelnetAsync(CancellationToken cancellationToken)
    {
        _chatTelnetPromptsSeen = 0;
        _chatTelnetCredentialsSent = false;
        LogInfo($"Battle.net connecting to {Config.BattlenetServer} (Chat protocol)...");
        await _chatTelnet.ConnectAsync(Config.BattlenetServer, Config.BattlenetPort, cancellationToken, BuildProxyOptions())
            .ConfigureAwait(false);
    }

    private async void OnChatTelnetConnected()
    {
        try
        {
            LogInfo("Battle.net Connected!");
            BncsConnected?.Invoke();
            await _chatTelnet.SendHandshakeAsync().ConfigureAwait(false);

            // A server that says nothing at all would otherwise leave this bot waiting forever:
            // war.pianka.io sends no banner and no prompt, just silence until it's given a
            // username and a password (probed live 2026-09-15). Every prompt-sending server seen
            // so far answers within a round trip, so a short wait tells the two apart — and once
            // a silent server has been identified, later reconnects skip the wait entirely, which
            // is what keeps rapid reconnect on one of these down to about a round trip.
            if (_chatTelnetServerIsSilent)
            {
                await SendChatTelnetCredentialsAsync().ConfigureAwait(false);
                return;
            }

            SafeFireAndForget(SendCredentialsIfNoPromptArrivesAsync(SilentServerWindow), "logging in to a server that sends no prompt");
        }
        catch (Exception ex)
        {
            LogError($"Error while starting the Chat-protocol handshake: {ex.Message}");
        }
    }

    private async Task HandleChatTelnetLineAsync(string line)
    {
        LogDebug($"Chat recv: {line}");
        if (line.Length == 0)
        {
            return;
        }

        if (_auth.LoggedOnToBncs)
        {
            var chatEvent = ChatTelnetEventParser.TryParse(line);
            if (chatEvent is not null)
            {
                await HandleChatEvent(chatEvent).ConfigureAwait(false);
            }

            return;
        }

        if (!_chatTelnetCredentialsSent)
        {
            var trimmed = line.TrimEnd();
            if (trimmed.EndsWith(':'))
            {
                _chatTelnetPromptsSeen++;
                if (_chatTelnetPromptsSeen == 1)
                {
                    await _chatTelnet.SendLineAsync(Config.Username).ConfigureAwait(false);
                }
                else if (_chatTelnetPromptsSeen == 2)
                {
                    _chatTelnetCredentialsSent = true;
                    await _chatTelnet.SendLineAsync(Config.Password).ConfigureAwait(false);
                }

                return;
            }

            // Some servers (atlas.bnetdocs.org confirmed live) never send field-specific
            // "Username:"/"Password:" prompts at all — just one instructional sentence
            // ("Enter your login name and password."), then silently expect the client to
            // send its username, then its password, each on its own line, with nothing
            // further in between — standard bare-telnet-login behavior. Trouble is, servers
            // that *do* send real prompts (the original sample this was built from) also say
            // basically the same sentence first, so seeing it isn't enough on its own to tell
            // the two apart. Resolved by not committing immediately: wait a beat for an actual
            // colon-terminated prompt to show up on its own; if none does, assume bare-telnet
            // and blind-send both lines. If a real prompt arrives first, _chatTelnetCredentialsSent
            // is already true by the time this fires and it's a no-op.
            var mentionsPassword = line.Contains("password", StringComparison.OrdinalIgnoreCase);
            var mentionsName = line.Contains("name", StringComparison.OrdinalIgnoreCase) ||
                                line.Contains("login", StringComparison.OrdinalIgnoreCase) ||
                                line.Contains("username", StringComparison.OrdinalIgnoreCase);
            if (mentionsPassword && mentionsName)
            {
                // The connect-time timer already covers this; this shorter one just gets a server
                // that announces itself this way logged in a little sooner.
                SafeFireAndForget(SendCredentialsIfNoPromptArrivesAsync(TimeSpan.FromMilliseconds(500)), "falling back to a bare-telnet login");
                return;
            }
        }

        // The rest of the login banner ("Connection from [...]", blank separator lines)
        // doesn't need a reply — just wait for the first numbered event line, whatever it is: normally
        // that's "2010 NAME <username>" confirming the logon, but treating *any*
        // recognized event ID as the login/logged-in boundary is more robust than
        // requiring 2010 specifically, in case a server skips straight to channel
        // events without it.
        var firstSpace = line.IndexOf(' ');
        if (firstSpace < 0 || !int.TryParse(line[..firstSpace], out var eventId) || eventId < 1000)
        {
            return;
        }

        await OnChatTelnetLoggedOnAsync(line, eventId).ConfigureAwait(false);
    }

    /// <summary>
    /// The Chat protocol's equivalent of OnLoggedOnAsync — no SID_ENTERCHAT/GETCHANNELLIST/
    /// FRIENDSLIST to send (none exist here), so this is just the same bookkeeping every other
    /// logon path does: mark the session on (which is also what lets chat be sent and stops a
    /// reconnect that's still counting down), then join the home channel with a "/join".
    /// </summary>
    private async Task OnChatTelnetLoggedOnAsync(string line, int eventId)
    {
        _auth.LoggedOnToBncs = true;
        _autoReconnectCts?.Cancel();
        _connectedAt = DateTimeOffset.UtcNow;

        if (eventId == ChatTelnetEventParser.NameConfirmationEventId)
        {
            var confirmedName = line.Split(' ', 3).ElementAtOrDefault(2) ?? Config.Username;
            LogInfo($"Logged on as: {confirmedName} using Chat protocol.");
        }
        else if (ChatTelnetEventParser.TryParse(line) is { } firstEvent)
        {
            await HandleChatEvent(firstEvent).ConfigureAwait(false);
        }

        if (!string.IsNullOrWhiteSpace(Config.HomeChannel))
        {
            await JoinHomeAsync().ConfigureAwait(false);
        }
    }

    /// <summary>Waits a beat for a real prompt to show up on its own; blind-sends the login if none does.</summary>
    private async Task SendCredentialsIfNoPromptArrivesAsync(TimeSpan window)
    {
        await Task.Delay(window).ConfigureAwait(false);
        if (_chatTelnetCredentialsSent || !_chatTelnet.IsConnected)
        {
            return;
        }

        _chatTelnetServerIsSilent = true;
        await SendChatTelnetCredentialsAsync().ConfigureAwait(false);
    }

    private async Task SendChatTelnetCredentialsAsync()
    {
        _chatTelnetCredentialsSent = true;
        LogDebug("Chat: sending the login without waiting for a prompt.");
        await _chatTelnet.SendLineAsync(Config.Username).ConfigureAwait(false);
        await _chatTelnet.SendLineAsync(Config.Password).ConfigureAwait(false);
    }
}
