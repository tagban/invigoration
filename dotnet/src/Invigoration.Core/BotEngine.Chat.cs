using Invigoration.Core.Chat;
using Invigoration.Core.Networking;

namespace Invigoration.Core;

/// <summary>
/// Battle.net/PVPGN's older plain-text, line-based "Chat" connection type —
/// selected via <see cref="Config"/>.ConnectionMode as an alternative to the
/// normal binary BNCS protocol, for PVPGN networks that still run it (e.g.
/// eurobattle.net). No BNLS, CD-key, or version-check involved: login is
/// just a two-prompt (username, then password) exchange — see
/// _chatTelnetPromptsSeen's remarks for why prompts are recognized by shape
/// rather than exact wording — then every subsequent line is a numbered
/// event (see <see cref="ChatTelnetEventParser"/>)
/// fed into the exact same <see cref="HandleChatEvent"/> pipeline the binary
/// protocol uses — the whole roster/rank/trivia/command-dispatch stack
/// downstream of that method doesn't need to know which transport it came
/// from. Confirmed against a live capture for the login handshake and the
/// USER/JOIN/TALK/CHANNEL events; the rest of the event-type mapping is a
/// best-effort extrapolation of the confirmed pattern (see
/// ChatTelnetEventParser's remarks) — expect to need live-testing fixes.
/// </summary>
public sealed partial class BotEngine
{
    private readonly ChatTelnetConnection _chatTelnet = new();
    private bool _chatTelnetLoggedIn;

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
            _chatTelnetLoggedIn = false;
            _friends.Clear();
            LogError($"Battle.net disconnected{(ex is null ? "." : $": {ex.Message}")}");
            BncsDisconnected?.Invoke(ex);
            MaybeScheduleAutoReconnect();
        };
    }

    private async Task ConnectChatTelnetAsync(CancellationToken cancellationToken)
    {
        _chatTelnetLoggedIn = false;
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
            await _chatTelnet.SendByteAsync(0x03).ConfigureAwait(false); // Connection type: Chat
            await _chatTelnet.SendByteAsync(0x04).ConfigureAwait(false); // Login sub-type
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

        if (_chatTelnetLoggedIn)
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
                SafeFireAndForget(SendCredentialsIfNoPromptArrivesAsync(), "falling back to a bare-telnet login");
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

        _chatTelnetLoggedIn = true;
        if (eventId == ChatTelnetEventParser.NameConfirmationEventId)
        {
            var confirmedName = line.Split(' ', 3).ElementAtOrDefault(2) ?? Config.Username;
            LogInfo($"Logged on as: {confirmedName} using Chat protocol.");
            return;
        }

        var firstEvent = ChatTelnetEventParser.TryParse(line);
        if (firstEvent is not null)
        {
            await HandleChatEvent(firstEvent).ConfigureAwait(false);
        }
    }

    private async Task SendCredentialsIfNoPromptArrivesAsync()
    {
        await Task.Delay(500).ConfigureAwait(false);
        if (_chatTelnetCredentialsSent)
        {
            return;
        }

        _chatTelnetCredentialsSent = true;
        await _chatTelnet.SendLineAsync(Config.Username).ConfigureAwait(false);
        await _chatTelnet.SendLineAsync(Config.Password).ConfigureAwait(false);
    }
}
