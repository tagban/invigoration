using System.Collections.ObjectModel;
using Avalonia.Media;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.Models;
using Invigoration.Core;
using Invigoration.Core.Chat;
using Invigoration.Core.Crypto;
using Invigoration.Core.Hotline;
using Invigoration.Core.Music;
using Invigoration.Core.Text;
using Invigoration.Core.Tracking;
using Invigoration.Core.Trivia;

namespace Invigoration.App.ViewModels;

/// <summary>
/// One live connection to a Hotline server — a sub-tab inside the outer Hotline group, the same
/// role BotTabViewModel plays inside a BotGroupTabViewModel. Owns a real HotlineTransactionClient;
/// every event it raises fires on the network receive-loop thread, so everything here marshals
/// onto the UI thread via Dispatcher.UIThread.Post before touching an ObservableCollection or
/// [ObservableProperty] (an ObservableCollection mutated off the UI thread throws).
/// </summary>
public sealed partial class HotlineSessionViewModel : ViewModelBase, IAsyncDisposable
{
    private readonly HotlineTabViewModel _parent;
    private readonly HotlineTransactionClient _client = new();
    private readonly HotlineConnectOptions _options;

    /// <summary>Whether the caller already gave us a real name (a saved profile's own Name, or the tracker listing's Name) — if so, that always wins over the server's own self-reported name once login completes.</summary>
    private readonly bool _hasExplicitDisplayName;

    public string Host { get; }
    public int Port { get; }

    /// <summary>The saved profile this session was connected from, if any — lets the tracker page show a "already connected" indicator/quick-open next to the matching saved server row instead of only a plain "Connect" button. Null for an ad-hoc connect straight from the tracker's server list.</summary>
    public string? ProfileId => _options.ProfileId;

    /// <summary>Off by default — gates "/trivia" on this server the same hard way BotEngine.Trivia.cs gates it for a Battle.net bot (see HotlineServerProfile.TriviaEnabled's remarks). Also gates last-seen tracking (MarkUserSeen) — no point writing to the tracking store for a server where trivia's never used.</summary>
    public bool TriviaEnabled => _options.TriviaEnabled;

    private readonly TriviaSession _triviaSession = new();
    private readonly TriviaEngine _triviaEngine;

    /// <summary>The same cosmetic text-transform toggles/prepend-postpend Battle.net bots have (BotSessionState) — session-scoped here too (not persisted), applied via the shared Chat.ChatTextEffects. See SendChatWithEffectsAsync.</summary>
    private bool _canadaMode;
    private bool _fuddMode;
    private bool _mooMode;
    private bool _leetSpeakMode;
    private string _prependText = "";
    private string _postpendText = "";

    [ObservableProperty]
    public partial string Title { get; set; }

    public IBrush HighlightBrush => HotlineTabViewModel.AccentBrush;
    public double HeaderFontSize => 13;
    public IBrush HeaderForeground => Brushes.White;

    [ObservableProperty]
    public partial bool HasUnread { get; set; }

    [ObservableProperty]
    public partial bool IsConnected { get; set; }

    [ObservableProperty]
    public partial string InputText { get; set; } = "";

    /// <summary>Non-null exactly when a server has prompted its agreement and it hasn't been accepted or dismissed yet — never auto-accepted unless the tracker's own AutoAcceptAgreement setting is on. See HotlineTransactionClient.AgreementReceived's remarks.</summary>
    [ObservableProperty]
    public partial string? PendingAgreementText { get; set; }

    /// <summary>Mirrors the tracker's own collapsed-by-default "Advanced" setting — off by default so the Copy Log button doesn't take up space for users who never need it. Reacts live to the tracker's ConfigChanged so toggling it in the Advanced section updates any already-open session immediately.</summary>
    public bool ShowCopyLogButton => _parent.Config.ShowCopyLogButton;

    public ObservableCollection<HotlineChatLine> Messages { get; } = [];
    public ObservableCollection<HotlineUserRowViewModel> Users { get; } = [];

    /// <summary>Discord users seen talking through this server's relay bot recently — kept separate from the real Users list, per explicit request. See PruneStaleGhosts and TryAppendDiscordRelayMessage.</summary>
    public ObservableCollection<HotlineGhostUserViewModel> DiscordUsers { get; } = [];

    private static readonly TimeSpan GhostExpiry = TimeSpan.FromMinutes(30);
    private readonly DispatcherTimer _ghostPruneTimer = new() { Interval = TimeSpan.FromMinutes(1) };

    public HotlineSessionViewModel(HotlineTabViewModel parent, HotlineConnectOptions options)
    {
        _parent = parent;
        _options = options;
        _triviaEngine = new TriviaEngine(new HotlineTriviaHost(this), () => _triviaSession);
        _ghostPruneTimer.Tick += (_, _) => PruneStaleGhosts();
        _ghostPruneTimer.Start();
        parent.ConfigChanged += () => OnPropertyChanged(nameof(ShowCopyLogButton));
        Host = options.Host;
        Port = options.Port;
        _hasExplicitDisplayName = !string.IsNullOrEmpty(options.DisplayName);
        Title = options.DisplayName is { Length: > 0 } ? options.DisplayName : $"{options.Host}:{options.Port}";
        _client.AutoAcceptAgreement = options.AutoAcceptAgreement;
        _client.Debug = parent.Config.Debug;

        _client.ChatMessageReceived += msg => Dispatcher.UIThread.Post(() => AppendChatMessage(msg));
        _client.ServerMessageReceived += msg => Dispatcher.UIThread.Post(() => AppendMessage($"* {msg}"));
        _client.UserListReceived += users => Dispatcher.UIThread.Post(() => ReplaceUsers(users));
        _client.UserChanged += user => Dispatcher.UIThread.Post(() => UpsertUser(user));
        _client.UserLeft += id => Dispatcher.UIThread.Post(() => RemoveUser(id));
        _client.AgreementReceived += text => Dispatcher.UIThread.Post(() => PendingAgreementText = text);
        _client.ProtocolError += ex => Dispatcher.UIThread.Post(() => AppendMessage($"* (internal) couldn't parse a message from the server: {ex.Message}"));
        _client.DisconnectMessageReceived += msg => Dispatcher.UIThread.Post(() => AppendMessage($"* Server says: {msg}"));
        _client.DebugLog += line => Dispatcher.UIThread.Post(() => AppendMessage($"* [debug] {line}"));
        _client.Disconnected += ex => Dispatcher.UIThread.Post(() =>
        {
            IsConnected = false;
            AppendMessage(ex is null ? "* Disconnected." : $"* Disconnected ({ex.GetType().Name}: {ex.Message}).");
        });

        _ = ConnectAsync(options.Login, options.Password, options.Nickname, options.IconId);
    }

    private async Task ConnectAsync(string login, string password, string nickname, ushort iconId)
    {
        AppendMessage($"* Connecting to {Host}:{Port}...");
        var ok = await _client.ConnectAndLoginAsync(Host, Port, login, password, nickname, iconId, _options.SendClientVersion ? _options.ClientVersion : null, _options.AdvertiseChatHistorySupport).ConfigureAwait(true);
        IsConnected = ok;
        AppendMessage(ok ? "* Connected." : "* Failed to connect or log in.");
        if (ok)
        {
            // A saved profile's or tracker listing's own name always wins — this is purely the
            // fallback for a session that started with neither (e.g. a manually-typed host).
            if (!_hasExplicitDisplayName && _client.ServerName is { Length: > 0 } serverName)
            {
                Title = serverName;
            }

            ReplaceUsers(_client.Users);

            // Two ways to show "where the conversation was last at" on connect, per explicit
            // request — mutually exclusive to avoid showing the same recent messages twice.
            // A server that speaks the "2.5" chat history extension has its own real, authoritative
            // record (can go back further than 10 lines) — prefer that. Only fall back to the
            // local RecentMessageStore cache (see its own remarks) for a pre-2.5 (1.2.3+) server,
            // which has no memory of the conversation at all.
            if (_client.SupportsChatHistory)
            {
                await LoadServerChatHistoryAsync().ConfigureAwait(true);
            }
            else
            {
                ShowLocalRecentMessages();
            }
        }
    }

    private string ServerTag => $"{Host}:{Port}";

    /// <summary>Pre-populates the chat log from the server's own persisted history (Get Chat History, 700) — see HotlineTransactionClient.GetChatHistoryAsync's remarks. Entries already arrive oldest-first, so they can be appended directly in the order received.</summary>
    private async Task LoadServerChatHistoryAsync()
    {
        var (entries, _) = await _client.GetChatHistoryAsync(limit: 20).ConfigureAwait(true);
        if (entries.Count == 0)
        {
            return;
        }

        AppendMessage("* ── chat history ──");
        foreach (var entry in entries)
        {
            AppendLine(RenderHistoryEntry(entry));
        }

        AppendMessage("* ── end of history ──");
    }

    private HotlineChatLine RenderHistoryEntry(HotlineChatHistoryEntry entry)
    {
        if (entry.IsDeleted)
        {
            return HotlineChatLine.Plain("[message removed]");
        }

        if (entry.IsServerMessage)
        {
            return HotlineChatLine.Plain($"* {entry.Message}");
        }

        // A historical sender's Admin status isn't carried by the entry itself — best-effort match
        // against whoever's currently online (same truncated-name convention as live chat, see
        // AppendChatMessage's remarks) so an admin's replayed lines still show in their rank color;
        // falls back to the default color for someone no longer connected.
        var matchName = entry.Nickname.Length > 13 ? entry.Nickname[..13] : entry.Nickname;
        var onlineUser = Users.FirstOrDefault(u => string.Equals(u.Name.Length > 13 ? u.Name[..13] : u.Name, matchName, StringComparison.Ordinal));
        var brush = onlineUser?.HighlightBrush
            ?? new SolidColorBrush(Color.Parse(onlineUser?.User.IsAdmin == true ? _parent.Config.AdminColorHex : _parent.Config.DefaultColorHex));

        var text = entry.IsAction ? $" {entry.Message}" : $":  {entry.Message}";
        return new HotlineChatLine(entry.Nickname, brush, text);
    }

    /// <summary>Pre-populates the chat log from the local RecentMessageStore cache — the fallback for a server that doesn't speak the "2.5" chat history extension. See RecentMessageStore's own remarks on why this exists separately.</summary>
    private void ShowLocalRecentMessages()
    {
        var recent = RecentMessageStore.GetRecent("Hotline", ServerTag);
        if (recent.Count == 0)
        {
            return;
        }

        AppendMessage("* ── last seen here ──");
        foreach (var message in recent)
        {
            AppendMessage(message.TimestampUtc is { } ts ? $"[{ts.ToLocalTime():t}] {message.Text}" : message.Text);
        }

        AppendMessage("* ── new messages below ──");
    }

    [RelayCommand]
    private async Task AcceptAgreement()
    {
        await _client.AcceptAgreementAsync().ConfigureAwait(true);
        PendingAgreementText = null;
    }

    /// <summary>The "x" close — per the user's own spec, dismissing the Agreement tab does NOT send Agreed; it just stops showing the prompt (the server presumably still limits an un-agreed session's access on its own).</summary>
    [RelayCommand]
    private void DismissAgreement() => PendingAgreementText = null;

    [RelayCommand]
    private async Task Send()
    {
        var text = InputText.Trim();
        if (text.Length == 0 || !IsConnected)
        {
            return;
        }

        InputText = "";

        if (text.StartsWith('/') && await TryHandleSlashCommandAsync(text[1..]).ConfigureAwait(true))
        {
            return;
        }

        await SendChatWithEffectsAsync(text).ConfigureAwait(true);
    }

    /// <summary>
    /// Applies the same Fudd/Canada/prepend/postpend transforms Battle.net bots have (see
    /// Chat.ChatTextEffects) to plain chat text — never to a "/"-prefixed line (a raw command or a
    /// "/me" trivia round message), matching BotEngine.SendChatCommandAsync's own isSlashCommand
    /// check. Public so HotlineTriviaHost can route every round message through the same path.
    /// </summary>
    public Task SendChatWithEffectsAsync(string text)
    {
        var outgoing = text.Length > 0 && text[0] == '/' ? text : ChatTextEffects.Apply(text, _fuddMode, _canadaMode, _prependText, _postpendText);
        return _client.SendChatAsync(outgoing);
    }

    /// <summary>Local-only diagnostic line (never sent to chat) — used by HotlineTriviaHost.LogParseErrors.</summary>
    public void AppendDebugMessage(string text) => AppendMessage($"* {text}");

    /// <summary>
    /// The same small set of protocol-agnostic commands Battle.net bots already have
    /// (BotEngine.Commands.cs's !ver/!nowplaying/!skip/etc.) — they only ever touch the global
    /// MusicPlayerRegistry and AppVersion, never anything Battle.net-specific, so they work
    /// identically here. Not routed through BotEngine itself (that's a whole Battle.net-connected
    /// bot, not something a Hotline session has or needs) — this is a small, deliberately separate
    /// implementation of just the commands that make sense outside a Battle.net context; Hotline
    /// doesn't get the richer clan/trivia/admin commands, which genuinely don't apply here.
    /// Returns false for anything it doesn't recognize, so the caller falls back to sending the
    /// original text as a literal chat message (matches a real Hotline client: an unrecognized
    /// "/word" is just chat, not an error).
    /// </summary>
    private async Task<bool> TryHandleSlashCommandAsync(string commandText)
    {
        var spaceIdx = commandText.IndexOf(' ');
        var command = (spaceIdx < 0 ? commandText : commandText[..spaceIdx]).ToLowerInvariant();
        var rest = spaceIdx < 0 ? "" : commandText[(spaceIdx + 1)..].Trim();

        switch (command)
        {
            case "ver":
                await _client.SendChatAsync($"/me is an Invigoration v{AppVersion.Current}").ConfigureAwait(true);
                return true;

            case "trivia":
                await _triviaEngine.HandleCommandAsync(rest, SendChatWithEffectsAsync).ConfigureAwait(true);
                return true;

            case "prepend":
            case "pre":
                _prependText = rest;
                AppendMessage($"* \"{rest}\" will be shown before each send.");
                return true;

            case "postpend":
            case "post":
                _postpendText = rest;
                AppendMessage($"* \"{rest}\" will be shown after each send.");
                return true;

            case "hex":
            case "h":
                await _client.SendChatAsync("£" + HexCodec.StrToHex(rest)).ConfigureAwait(true);
                return true;

            case "invigencrypt":
            case "encrypt":
            case "ie":
            case "i":
                await _client.SendChatAsync(InvigCipher.Encrypt(rest + "-")).ConfigureAwait(true);
                return true;

            case "canada":
                _canadaMode = !_canadaMode;
                await SendChatWithEffectsAsync($"Canada Mode {(_canadaMode ? "enabled" : "disabled")}.").ConfigureAwait(true);
                return true;

            case "fudd":
                _fuddMode = !_fuddMode;
                await SendChatWithEffectsAsync($"Elmer Fudd mode {(_fuddMode ? "enabled" : "disabled")}.").ConfigureAwait(true);
                return true;

            case "moo":
                _mooMode = !_mooMode;
                await SendChatWithEffectsAsync(_mooMode ? "Moooooooooooooooo mode engaged!" : "Cows are off...").ConfigureAwait(true);
                return true;

            case "leetspeak":
                _leetSpeakMode = !_leetSpeakMode;
                await SendChatWithEffectsAsync($"Leet speak {(_leetSpeakMode ? "enabled" : "disabled")}.").ConfigureAwait(true);
                return true;

            case "debug":
                _client.Debug = !_client.Debug;
                AppendMessage($"* Debug mode {(_client.Debug ? "enabled" : "disabled")}.");
                return true;

            case "setusername":
                if (rest.Length == 0)
                {
                    AppendMessage("* Usage: /setusername <new nickname>");
                    return true;
                }

                await _client.ChangeUserInfoAsync(rest).ConfigureAwait(true);
                AppendMessage($"* Nickname changed to \"{rest}\".");
                return true;

            case "nowplaying":
            case "np":
            case "music":
                await SendMusicReplyAsync(await GetNowPlayingReplyAsync().ConfigureAwait(true)).ConfigureAwait(true);
                return true;

            case "skip":
            case "next":
                await RunMusicCommandAsync(c => c.SkipAsync(), "Skipped.", "Couldn't skip — is a track playing?").ConfigureAwait(true);
                return true;

            case "pause":
            case "play":
            case "stop":
                await RunMusicCommandAsync(c => c.PlayPauseAsync(), "Toggled play/pause.", "Couldn't toggle play/pause — is the music player open?").ConfigureAwait(true);
                return true;

            case "thumbsup":
                await RunMusicCommandAsync(c => c.ThumbsUpAsync(), "Liked it.", "Couldn't like the current track — make sure you're signed in to the music player.", c => c.SupportsThumbsUp).ConfigureAwait(true);
                return true;

            case "thumbsdown":
                await RunMusicCommandAsync(c => c.ThumbsDownAsync(), "Disliked it.", "Couldn't dislike the current track — make sure you're signed in to the music player.", c => c.SupportsThumbsDown).ConfigureAwait(true);
                return true;

            default:
                return false;
        }
    }

    private static async Task<string> GetNowPlayingReplyAsync()
    {
        if (MusicPlayerRegistry.Controller is not { } controller)
        {
            return "Music player isn't open.";
        }

        var nowPlaying = await controller.GetNowPlayingAsync().ConfigureAwait(false);
        return nowPlaying is null
            ? "Nothing seems to be playing."
            : $"/me is now playing {nowPlaying.Title} - by {nowPlaying.Artist}{(string.IsNullOrEmpty(nowPlaying.Service) ? "" : $" on {nowPlaying.Service}")}.";
    }

    private async Task RunMusicCommandAsync(Func<IMusicPlayerController, Task<bool>> action, string successText, string failureText, Func<IMusicPlayerController, bool>? isSupported = null)
    {
        if (MusicPlayerRegistry.Controller is not { } controller)
        {
            await SendMusicReplyAsync("Music player isn't open.").ConfigureAwait(true);
            return;
        }

        if (isSupported is not null && !isSupported(controller))
        {
            return;
        }

        await SendMusicReplyAsync(await action(controller).ConfigureAwait(true) ? successText : failureText).ConfigureAwait(true);
    }

    private Task SendMusicReplyAsync(string text) => SendChatWithEffectsAsync(text);

    [RelayCommand]
    private async Task Disconnect()
    {
        await _client.DisposeAsync().ConfigureAwait(true);
        _parent.CloseSession(this);
    }

    private void AppendMessage(string line) => AppendLine(HotlineChatLine.Plain(line));

    /// <summary>
    /// Colorizes the sender's name by rank (Admin vs everyone else — see
    /// HotlineTrackerConfig.AdminColorHex's remarks) if one can be found. The protocol gives no
    /// structured sender field for a chat broadcast — Mobius's own source confirms the server
    /// pre-formats the whole line as one string, "\r{name,13}:  {message}" (or "\r*** {name}
    /// {message}" for /me) — so this matches the raw text against the *live* Users list rather
    /// than trying to positionally parse a format that can vary by server/version, and skips
    /// coloring entirely (falls back to a plain line) if no known name matches.
    ///
    /// Matches against a name *truncated to 13 characters*, not the full stored name — Hotline's
    /// own "%13.13s" formatting truncates (not just pads) a longer username in the chat line
    /// itself, so a straight full-name match would never fire for anyone with a name over 13
    /// characters. The colored display still uses the real, full name from the Users list.
    /// </summary>
    private void AppendChatMessage(string rawText)
    {
        var trimmed = rawText.TrimStart('\r');

        if (!string.IsNullOrEmpty(_options.DiscordRelayUsername) && TryAppendDiscordRelayMessage(trimmed))
        {
            return;
        }

        foreach (var user in Users)
        {
            var matchName = user.Name.Length > 13 ? user.Name[..13] : user.Name;

            var normalMarker = $"{matchName}:  ";
            var idx = trimmed.IndexOf(normalMarker, StringComparison.Ordinal);
            if (idx >= 0 && trimmed[..idx].Trim().Length == 0)
            {
                // Only a genuine "Talk" line (not a "/me" Emote below) is tracked/matched as a
                // trivia answer — mirrors BotEngine.Bncs.cs's own ChatEventType.Talk-only handling.
                if (TriviaEnabled)
                {
                    var messageText = trimmed[(idx + normalMarker.Length)..];
                    HandleUserChatLine(user, messageText);
                }

                // Skip past the (possibly truncated) name here — the colored Username run already
                // renders the real name, so the remainder must start at ":  message".
                AppendColorizedLine(user, trimmed[(idx + matchName.Length)..]);
                return;
            }

            var meMarker = $"*** {matchName} ";
            if (trimmed.StartsWith(meMarker, StringComparison.Ordinal))
            {
                if (TriviaEnabled)
                {
                    MarkUserSeen(user);
                }

                AppendColorizedLine(user, trimmed[(meMarker.Length - 1)..]);
                return;
            }
        }

        AppendMessage(rawText);
    }

    /// <summary>
    /// Marks the sender as seen (Tracking.ProtocolUserTrackingStore) and, if a trivia round is
    /// currently waiting on an answer, checks the message against the current question — mirrors
    /// BotEngine.Bncs.cs's own ChatEventType.Talk handling (Clan.ClanRosterStore.RecordSeen +
    /// _trivia.TryMatchAnswer). Gated behind TriviaEnabled entirely (see its remarks) — no point
    /// writing to the tracking store, or even checking a question that can never be running, for a
    /// server where trivia's off.
    /// </summary>
    private void HandleUserChatLine(HotlineUserRowViewModel user, string messageText)
    {
        MarkUserSeen(user);

        if (_triviaSession.IsEnabled && _triviaSession.TryMatchAnswer(messageText, out var matchedAnswer))
        {
            _triviaSession.PendingAnswer = (user.Name, matchedAnswer, $"Hotline: {Title}");
        }
    }

    private void MarkUserSeen(HotlineUserRowViewModel user) =>
        ProtocolUserTrackingStore.MarkSeen(user.Name, "Hotline", $"{Host}:{Port}");

    /// <summary>
    /// Parses this server's Discord relay convention — confirmed live across two real servers
    /// using genuinely different shapes: bigredh sends "Discord | {user}: {message}" from a
    /// Hotline account named "Relay"; MacDomain sends bare "{user}: {message}" from an account
    /// named "Discord". Both the relay account's own name and the optional prefix are per-tracker
    /// settings (HotlineTrackerConfig.DiscordRelayUsername/DiscordRelayPrefix) rather than
    /// hardcoded, since there's no way to auto-detect either from the protocol itself. Declutters
    /// the chat log by rendering the real Discord sender directly (with the Discord icon) instead
    /// of a wall of identical "Relay:" lines, and tracks them as a "ghost" user list entry.
    /// Returns false (falls through to normal rendering) if the line doesn't actually match — a
    /// relay bot can still send other messages (join/leave notices, errors) that don't fit the
    /// "{user}: {message}" shape.
    /// </summary>
    private bool TryAppendDiscordRelayMessage(string trimmed)
    {
        var relayName = _options.DiscordRelayUsername;
        var matchName = relayName.Length > 13 ? relayName[..13] : relayName;
        var marker = $"{matchName}:  ";
        var idx = trimmed.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (idx < 0 || trimmed[..idx].Trim().Length != 0)
        {
            return false;
        }

        var content = trimmed[(idx + marker.Length)..];
        var prefix = _options.DiscordRelayPrefix;
        if (!string.IsNullOrEmpty(prefix))
        {
            if (!content.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            content = content[prefix.Length..];
        }

        var colonIdx = content.IndexOf(": ", StringComparison.Ordinal);
        if (colonIdx <= 0)
        {
            return false;
        }

        var discordUser = content[..colonIdx];
        var ghost = UpsertDiscordGhost(discordUser);
        var brush = ghost.HighlightBrush ?? new SolidColorBrush(Color.Parse(_parent.Config.DefaultColorHex));
        var line = new HotlineChatLine(discordUser, brush, content[colonIdx..], GameIconLoader.Get("discord-relay"));
        AppendLine(line);
        RecentMessageStore.Append("Hotline", ServerTag, line.FullText, DateTimeOffset.UtcNow);
        return true;
    }

    private HotlineGhostUserViewModel UpsertDiscordGhost(string name)
    {
        PruneStaleGhosts();
        var existing = DiscordUsers.FirstOrDefault(g => string.Equals(g.Name, name, StringComparison.OrdinalIgnoreCase));
        if (existing is not null)
        {
            existing.LastSeen = DateTimeOffset.UtcNow;
            return existing;
        }

        var ghost = new HotlineGhostUserViewModel(_parent, name);
        DiscordUsers.Add(ghost);
        return ghost;
    }

    private void PruneStaleGhosts()
    {
        var cutoff = DateTimeOffset.UtcNow - GhostExpiry;
        for (var i = DiscordUsers.Count - 1; i >= 0; i--)
        {
            if (DiscordUsers[i].LastSeen < cutoff)
            {
                DiscordUsers.RemoveAt(i);
            }
        }
    }

    private void AppendColorizedLine(HotlineUserRowViewModel user, string remainder)
    {
        // A per-user highlight override always wins over the rank-based color — the whole point
        // is making a specific person's messages stand out regardless of their Admin status.
        var brush = user.HighlightBrush ?? new SolidColorBrush(Color.Parse(user.User.IsAdmin ? _parent.Config.AdminColorHex : _parent.Config.DefaultColorHex));
        var line = new HotlineChatLine(user.Name, brush, remainder);
        AppendLine(line);
        RecentMessageStore.Append("Hotline", ServerTag, line.FullText, DateTimeOffset.UtcNow);
    }

    private void AppendLine(HotlineChatLine line)
    {
        Messages.Add(line);
        if (_parent.SelectedItem != this)
        {
            HasUnread = true;
        }
    }

    private void ReplaceUsers(IReadOnlyList<HotlineUser> users)
    {
        Users.Clear();
        foreach (var user in users)
        {
            Users.Add(new HotlineUserRowViewModel(_parent, user));
        }
    }

    private void UpsertUser(HotlineUser user)
    {
        var existing = Users.FirstOrDefault(u => u.UserId == user.UserId);
        if (existing is not null)
        {
            Users.Remove(existing);
        }

        Users.Add(new HotlineUserRowViewModel(_parent, user));
    }

    private void RemoveUser(ushort userId)
    {
        var existing = Users.FirstOrDefault(u => u.UserId == userId);
        if (existing is not null)
        {
            Users.Remove(existing);
        }
    }

    public ValueTask DisposeAsync()
    {
        _ghostPruneTimer.Stop();
        return _client.DisposeAsync();
    }
}
