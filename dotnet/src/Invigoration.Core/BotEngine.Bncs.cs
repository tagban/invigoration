using Invigoration.Core.Auth;
using Invigoration.Core.Chat;
using Invigoration.Core.Networking;
using Invigoration.Core.Protocol;

namespace Invigoration.Core;

/// <summary>Port of modBNET.bas's ParseBnet. Frames are passed whole (header included) to keep offsets 1:1 with the original.</summary>
public sealed partial class BotEngine
{
    private Task HandleBncsPacket(byte[] frame)
    {
        var id = (BncsPacketId)BncsConnection.GetPacketId(frame);
        LogDebug($"BNCS recv 0x{(byte)id:X2} ({id}), {frame.Length} bytes: {ToHexDump(frame)}");

        return id switch
        {
            BncsPacketId.SID_PING => HandlePingAsync(),
            BncsPacketId.SID_AUTH_INFO => HandleAuthInfoAsync(frame),
            BncsPacketId.SID_AUTH_CHECK => HandleAuthCheckAsync(frame),
            BncsPacketId.SID_AUTH_ACCOUNTCREATE => HandleAuthAccountCreateAsync(frame),
            BncsPacketId.SID_AUTH_ACCOUNTLOGON => HandleAuthAccountLogonAsync(frame),
            BncsPacketId.SID_AUTH_ACCOUNTLOGONPROOF => HandleAuthAccountLogonProofAsync(frame),
            BncsPacketId.SID_LOGONRESPONSE2 => HandleLogonResponse2Async(frame),
            BncsPacketId.SID_QUERYREALMS => HandleQueryRealmsReplyAsync(),
            BncsPacketId.SID_LOGONREALMEX => HandleLogonRealmExAsync(frame),
            BncsPacketId.SID_ENTERCHAT => HandleEnterChat(frame),
            BncsPacketId.SID_GETCHANNELLIST => HandleGetChannelList(frame),
            BncsPacketId.SID_CHATEVENT => HandleBncsChatEventFrame(frame),
            BncsPacketId.SID_NEWS_INFO => HandleNewsInfo(frame),
            BncsPacketId.SID_REQUIREDWORK => HandleRequiredWork(),
            BncsPacketId.SID_CREATEACCOUNT => HandleLegacyCreateAccountReplyAsync(),
            BncsPacketId.SID_SETEMAIL => HandleSetEmail(),
            BncsPacketId.SID_FRIENDSLIST => HandleFriendsList(frame),
            BncsPacketId.SID_FRIENDSUPDATE => HandleFriendsUpdate(frame),
            BncsPacketId.SID_FRIENDSADD => HandleFriendsAdd(frame),
            BncsPacketId.SID_FRIENDSREMOVE => HandleFriendsRemove(frame),
            BncsPacketId.SID_FRIENDSPOSITION => HandleFriendsPosition(frame),
            _ => Task.CompletedTask,
        };
    }

    private Task HandlePingAsync()
    {
        if (Config.NegPing || Config.ZeroPing)
        {
            // NegPing: never respond, server shows a negative/placeholder ping.
            // ZeroPing: the one fabricated response sent at connect stands; not
            // responding again means the server can't recalculate it.
            return Task.CompletedTask;
        }

        // Ported as-is from Send0x25 in modBNET.bas: echoes a hardcoded 0
        // rather than the cookie actually received. That's what shipped and
        // was confirmed working against PVPGN/Atlas, so it's kept rather than
        // "corrected" to a real cookie echo.
        return SendBncsAsync(new PacketWriter().WriteDword(0), BncsPacketId.SID_PING);
    }

    private async Task HandleAuthInfoAsync(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        reader.Skip(4); // Logon Type
        _auth.ServerToken = reader.ReadDword();
        reader.ReadDword(); // UDP Value, unused
        var mpqFileTime = reader.ReadFileTime();
        var mpqFileName = reader.ReadNTString();
        var checkRevisionFormula = reader.ReadNTString();

        var writer = new PacketWriter()
            .WriteDword(BncsProduct.GetBnlsProductByte(Config.Product) ?? 0)
            .WriteDword(0) // flags
            .WriteDword(1) // cookie
            .WriteDword(mpqFileTime.Low)
            .WriteDword(mpqFileTime.High)
            .WriteNTString(mpqFileName)
            .WriteNTString(checkRevisionFormula);
        await SendBnlsAsync(writer, BnlsPacketId.BNLS_VERSIONCHECKEX2).ConfigureAwait(false);
    }

    private async Task HandleAuthCheckAsync(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        var status = reader.ReadWord();

        switch (status)
        {
            case 0x0000:
                LogInfo("Version check + CD-key check passed.");
                await OnAuthCheckPassedAsync().ConfigureAwait(false);
                break;
            case 0x0100:
                LogError("Game version out of date.");
                break;
            case 0x0101:
                LogError("Invalid game version. Check your BNLS/JBLS server and try another one.");
                break;
            case 0x0102:
                LogError("Game version needs to be downgraded.");
                break;
            case 0x0200:
                LogError("Invalid CD key.");
                break;
            case 0x0201:
                LogError("CD key is in use.");
                break;
            case 0x0202:
                LogError("Current CD key is banned from Battle.net.");
                break;
            case 0x0203:
                LogError("Incorrect CD key for this product. Please check your key and/or game.");
                break;
        }
    }

    private async Task OnAuthCheckPassedAsync()
    {
        if (BncsProduct.UsesNewLoginSystem(Config.Product))
        {
            var writer = new PacketWriter().WriteNTString(Config.Username).WriteNTString(Config.Password);
            await SendBnlsAsync(writer, BnlsPacketId.BNLS_LOGONCHALLENGE).ConfigureAwait(false);
            return;
        }

        await SendBncsAsync(
            new PacketWriter().WriteAscii(Config.UseUdp ? "bnet" : "tenb"),
            BncsPacketId.SID_UDPPINGRESPONSE).ConfigureAwait(false);
        await SendBncsAsync(new PacketWriter(), BncsPacketId.SID_GETICONDATA).ConfigureAwait(false);

        _auth.HashPurpose = _auth.ChangePasswordRequested ? HashPurpose.ChangePassword : HashPurpose.AccountLogon;
        _auth.ChangePasswordRequested = false;
        _auth.HashStage = 0;
        await SendPasswordHashRequestAsync(Config.Password).ConfigureAwait(false);
    }

    private async Task HandleAuthAccountCreateAsync(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        // The status field is a DWORD, not a WORD — ReadWord() only consumed the low 2 of its 4
        // bytes. Coincidentally read the same numeric value for every status seen so far (the
        // high word is always zero), so this wasn't corrupting anything downstream, but it's the
        // wrong width regardless and worth fixing outright.
        var status = reader.ReadDword();
        if (status == 0)
        {
            LogInfo("Account created! Connecting with your new account...");
            var writer = new PacketWriter().WriteNTString(Config.Username).WriteNTString(Config.Password);
            await SendBnlsAsync(writer, BnlsPacketId.BNLS_LOGONCHALLENGE).ConfigureAwait(false);
            return;
        }

        // This branch never used to run at all — a failed create just silently went nowhere,
        // looking indistinguishable from a hung connection. Known official codes per bnetdocs;
        // anything else (this server returned 8, which isn't one of them) is very likely a
        // PVPGN-specific extension — reported as-is rather than guessed at.
        var reason = status switch
        {
            2 => "the name contains invalid characters",
            3 => "the name contains a banned word",
            4 => "an account with this name already exists",
            6 => "the name doesn't contain enough alphanumeric characters",
            7 => "the name contains too many characters of the same type in a row",
            _ => $"server-specific reason (status {status})",
        };
        LogError($"Could not create account \"{Config.Username}\": {reason}.");
    }

    private async Task HandleAuthAccountLogonAsync(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        var status = reader.ReadDword();
        if (status == 1)
        {
            LogWarning("Account doesn't exist, attempting to create it.");
            var writer = new PacketWriter().WriteNTString(Config.Username).WriteNTString(Config.Password);
            await SendBnlsAsync(writer, BnlsPacketId.BNLS_CREATEACCOUNT).ConfigureAwait(false);
            return;
        }

        var saltAndServerKey = reader.ReadRaw(64); // Salt(32) + server public key B(32)
        await SendBnlsAsync(new PacketWriter().WriteBytes(saltAndServerKey), BnlsPacketId.BNLS_LOGONPROOF)
            .ConfigureAwait(false);
    }

    private async Task HandleAuthAccountLogonProofAsync(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        switch (reader.ReadWord())
        {
            case 0x0000:
            case 0x000E: // success; server also wants an email address on file, ignored (matches original)
                await OnLoggedOnAsync().ConfigureAwait(false);
                break;
            case 0x0002:
                LogError("Battle.net logon failed: incorrect password.");
                break;
            case 0x0006:
                LogError("This account was closed or banned.");
                break;
        }
    }

    private async Task HandleLogonResponse2Async(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        switch (reader.ReadByte())
        {
            case 0x01:
                LogError("Battle.net logon failed!");
                if (!_auth.AttemptedAccountCreate)
                {
                    _auth.AttemptedAccountCreate = true;
                    _auth.HashPurpose = HashPurpose.AccountCreate;
                    _auth.HashStage = 0;
                    await SendPasswordHashRequestAsync(Config.Password).ConfigureAwait(false);
                }

                break;
            case 0x02:
                LogError("Battle.net logon failed, due to incorrect password.");
                break;
            case 0x06:
                LogError("This account was closed or banned.");
                break;
            case 0x00:
                if (Config.Realm.Length > 0 &&
                    _auth.WantsRealmLogon &&
                    Config.Product is BncsProduct.DiabloII or BncsProduct.DiabloIILoD)
                {
                    await OnLoggedOnAsync(enterChat: false).ConfigureAwait(false);
                    await SendBncsAsync(
                        new PacketWriter().WriteDword(0).WriteDword(0).WriteByte(0),
                        BncsPacketId.SID_QUERYREALMS).ConfigureAwait(false);
                }
                else
                {
                    await OnLoggedOnAsync(joinHomeChannelDirectly: true).ConfigureAwait(false);
                }

                break;
        }
    }

    /// <summary>
    /// Enters chat and joins a channel after a successful logon.
    /// <paramref name="joinHomeChannelDirectly"/> distinguishes the two
    /// original code paths: the old login system (SID_LOGONRESPONSE2, used
    /// by D2/W2BN) joins the configured home channel directly (flags=2); the
    /// new NLS login system (SID_AUTH_ACCOUNTLOGONPROOF, used by WC3/TFT)
    /// joins the "L" pseudo-channel (flags=1, "last channel") instead. These
    /// were previously incorrectly unified, which sent every product into
    /// "L" regardless of the configured home channel.
    /// </summary>
    private async Task OnLoggedOnAsync(bool enterChat = true, bool joinHomeChannelDirectly = false)
    {
        LogInfo("Battle.net Logon Passed!");
        _auth.LoggedOnToBncs = true;
        _auth.AttemptedAccountCreate = false;
        _connectedAt = DateTimeOffset.UtcNow;

        if (!enterChat)
        {
            return;
        }

        await SendBncsAsync(new PacketWriter().WriteNTString(Config.Username).WriteByte(0), BncsPacketId.SID_ENTERCHAT)
            .ConfigureAwait(false);
        await SendBncsAsync(new PacketWriter().WriteAscii(Config.Product), BncsPacketId.SID_GETCHANNELLIST)
            .ConfigureAwait(false);

        if (BncsProduct.SupportsFriendsList(Config.Product))
        {
            await RequestFriendsListAsync().ConfigureAwait(false);
        }

        var joinWriter = joinHomeChannelDirectly
            ? new PacketWriter().WriteDword(2).WriteNTString(Config.HomeChannel)
            : new PacketWriter().WriteDword(1).WriteNTString("L");
        await SendBncsAsync(joinWriter, BncsPacketId.SID_JOINCHANNEL).ConfigureAwait(false);
    }

    /// <summary>
    /// SID_QUERYREALMS reply. The VB6 original hard-coded the literal ASCII
    /// text "password" here instead of the real account password — a bug
    /// that would make realm logon fail for every user. Fixed to hash the
    /// actual password (D2 realm logon reuses the main account password).
    /// </summary>
    private Task HandleQueryRealmsReplyAsync()
    {
        _auth.HashPurpose = HashPurpose.RealmLogon;
        _auth.HashStage = 0;
        return SendPasswordHashRequestAsync(Config.Password);
    }

    private async Task HandleLogonRealmExAsync(byte[] frame)
    {
        // Offset matches modBNET.bas's derivation (P1=payload[0..15], the
        // "Server" chunk at payload[12..19], with the IP itself 4 bytes into
        // that chunk at payload[16..19]) rather than a from-scratch reading
        // of the MCP chunk layout, since that's what's confirmed working
        // against PVPGN.
        var reader = BncsConnection.GetPayloadReader(frame);
        reader.Skip(16);
        var realmServerIp = reader.ReadRaw(4);
        var host = $"{realmServerIp[0]}.{realmServerIp[1]}.{realmServerIp[2]}.{realmServerIp[3]}";
        LogInfo($"Current realm server: {host}");

        _realm.Close();
        await _realm.ConnectAsync(host, RealmPort, proxy: BuildProxyOptions()).ConfigureAwait(false);
    }

    private Task HandleEnterChat(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        var uniqueName = reader.ReadNTString();
        var statString = reader.ReadNTString();
        LogInfo($"Logged on as: {uniqueName} using {BncsProduct.GetDisplayName(Config.Product)}.");
        return Task.CompletedTask;
    }

    private Task HandleGetChannelList(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        var channels = new List<string>();
        while (reader.Remaining > 1)
        {
            var name = reader.ReadNTString();
            if (name.Length == 0)
            {
                break;
            }

            channels.Add(name);
        }

        ChannelListReceived?.Invoke(channels);
        return Task.CompletedTask;
    }

    /// <summary>Per-engine cache of each user's most recently seen 4-char product code, from ShowUser/Join/UserFlags statstrings — consulted when they later talk, so RecordSeen can stamp LastSeenProduct at the same time as LastSeenUtc.</summary>
    private readonly Dictionary<string, string> _lastKnownProduct = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Per-user rolling window of recent Join/Leave timestamps, for Config.HideJoinLeaveSpamEnabled — see IsJoinLeaveNoisy.</summary>
    private readonly Dictionary<string, List<DateTime>> _recentJoinLeaveTimestamps = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Parses a raw binary SID_CHATEVENT frame, then hands off to the shared (transport-agnostic) HandleChatEvent below — the Chat-protocol connection feeds the same method from its own line parser (see BotEngine.Chat.cs).</summary>
    private Task HandleBncsChatEventFrame(byte[] frame) => HandleChatEvent(ChatEventParser.Parse(frame));

    /// <summary>
    /// A human-readable "where this came from" for a trivia winner announcement — the
    /// motivating case is a TriviaGroup round spanning several bots/products at once, where
    /// it's genuinely ambiguous which linked channel (or Discord) a given answer arrived from.
    /// </summary>
    private string DescribeChatSource(ChatEvent chatEvent)
    {
        if (chatEvent.Origin == ChatEventOrigin.Discord)
        {
            return "Discord";
        }

        if (chatEvent.ChannelIndex is { } idx && _sc2Channels.TryGetValue(idx, out var session))
        {
            return $"StarCraft II - {session.Channel.Name}";
        }

        var channel = string.IsNullOrEmpty(_session.CurrentChannelName) ? "its channel" : _session.CurrentChannelName;
        return $"{Config.BattlenetServer} - {channel}";
    }

    private async Task HandleChatEvent(ChatEvent chatEvent)
    {
        // Display-only filter: everything below (roster, rank behaviors, counts, trivia,
        // commands) still runs exactly as normal regardless of this — only whether the
        // event also gets written to the visible chat log depends on it.
        var isHiddenJoinLeaveSpam = chatEvent.Type is ChatEventType.Join or ChatEventType.Leave &&
                                     (Config.SuppressJoinLeaveNotifications ||
                                      (Config.HideJoinLeaveSpamEnabled && IsJoinLeaveNoisy(chatEvent.Username)));
        if (!isHiddenJoinLeaveSpam)
        {
            ChatMessage?.Invoke(chatEvent);
        }

        // "[CHAT]" (and similar bracketed tags) is what Chat-protocol-only participants report
        // instead of a real 4-char BNCS product code — not a product, so don't track it as one.
        if (chatEvent.Type is ChatEventType.ShowUser or ChatEventType.Join or ChatEventType.UserFlags &&
            chatEvent.Text.Length >= 4 && chatEvent.Text[0] != '[')
        {
            var product = chatEvent.Text[..4];
            _lastKnownProduct[chatEvent.Username] = product;
            Clan.ClanRosterStore.RecordProductSeen(chatEvent.Username, product, Config.BattlenetServer);
        }

        if (chatEvent.Type is ChatEventType.ShowUser or ChatEventType.Join)
        {
            await ApplyRankBehaviorsAsync(chatEvent.Username).ConfigureAwait(false);
        }

        // Counts reset whenever the bot (re)joins a channel, then track that channel's own
        // activity from there — matches the VB6 original (ChatBot_OnChannel/OnJoin/OnInfo).
        if (chatEvent.Type == ChatEventType.Channel)
        {
            _session.BanCount = 0;
            _session.KickCount = 0;
            _session.JoinCount = 0;
            _session.CurrentChannelName = chatEvent.Text;
        }

        if (chatEvent.Type == ChatEventType.Join)
        {
            _session.JoinCount++;
        }

        if (chatEvent.Type == ChatEventType.Info)
        {
            if (chatEvent.Text.Contains("was kicked out of the channel by", StringComparison.OrdinalIgnoreCase))
            {
                _session.KickCount++;
            }
            else if (chatEvent.Text.Contains("was banned by", StringComparison.OrdinalIgnoreCase))
            {
                _session.BanCount++;
            }
        }

        if (chatEvent.Type is ChatEventType.Talk or ChatEventType.Emote or ChatEventType.Whisper)
        {
            var defaultRank = Config.ClanFeatureEnabled ? Config.DefaultRank : null;
            _lastKnownProduct.TryGetValue(chatEvent.Username, out var product);
            Clan.ClanRosterStore.RecordSeen(chatEvent.Username, defaultRank, product, Config.BattlenetServer);
        }

        // ChannelIndex is null for BNCS/Chat-Telnet (single-channel by protocol, so always a
        // match) and for a whisper on any product (not channel-scoped). For a Stimpak-backed
        // (SC2/SC:R/WC3:R) bot, this keeps an answer typed in one joined channel from
        // resolving a trivia question the bot posed in a different one.
        if (chatEvent.Type == ChatEventType.Talk && _trivia.IsEnabled && !IsBannedUser(chatEvent.Username) &&
            (chatEvent.ChannelIndex is null || chatEvent.ChannelIndex == _sc2TriviaChannelIndex) &&
            _trivia.TryMatchAnswer(chatEvent.Text, out var matchedAnswer))
        {
            _trivia.PendingAnswer = (chatEvent.Username, matchedAnswer, DescribeChatSource(chatEvent));
        }

        if (chatEvent.Type is ChatEventType.Talk or ChatEventType.Whisper)
        {
            if (chatEvent.Type == ChatEventType.Whisper)
            {
                _session.LastWhisperFromUser = chatEvent.Username;
                _session.LastWhisperFromText = chatEvent.Text;
            }

            await HandleCommandAsync(
                chatEvent.Username, chatEvent.Text, isWhisper: chatEvent.Type == ChatEventType.Whisper, chatEvent.ChannelIndex)
                .ConfigureAwait(false);
        }

        if (chatEvent.Type == ChatEventType.Info && chatEvent.Text.Trim().Equals("No one hears you.", StringComparison.OrdinalIgnoreCase))
        {
            await RecoverFromChannelDesyncAsync().ConfigureAwait(false);
        }
    }

    /// <summary>
    /// True once this user has racked up more than Config.HideJoinLeaveSpamThreshold
    /// Join/Leave events within the last Config.HideJoinLeaveSpamWindowSeconds —
    /// records the current event as part of the same check, so this is only
    /// ever called once per event (not a separate peek-then-record step).
    /// Self-correcting: once a noisy user's rate drops (their older
    /// timestamps age out of the window), a later join/leave stops counting
    /// as noisy again without needing anything to explicitly reset it.
    /// </summary>
    private bool IsJoinLeaveNoisy(string username)
    {
        var now = DateTime.UtcNow;
        var window = TimeSpan.FromSeconds(Math.Max(1, Config.HideJoinLeaveSpamWindowSeconds));

        if (!_recentJoinLeaveTimestamps.TryGetValue(username, out var timestamps))
        {
            timestamps = [];
            _recentJoinLeaveTimestamps[username] = timestamps;
        }

        timestamps.RemoveAll(t => now - t > window);
        timestamps.Add(now);
        return timestamps.Count > Math.Max(0, Config.HideJoinLeaveSpamThreshold);
    }

    /// <summary>
    /// Applies whatever automated behaviors this tracked member's current
    /// rank carries (see ClanRank) — a welcome whisper, and/or an automatic
    /// kick/ban for flagging troublemakers without needing to watch for them
    /// manually. Fires on ShowUser (the bot's own initial channel roster —
    /// so someone already in the channel when the bot connects is caught
    /// too) and Join, not on every UserFlags update, so this can't re-fire
    /// repeatedly for someone who's just sitting in the channel. A no-op
    /// when clan management is off for this bot, the speaker isn't tracked,
    /// or their rank isn't one of the predefined ClanRankStore ranks (a
    /// legacy free-text rank has no behaviors to apply).
    /// </summary>
    private async Task ApplyRankBehaviorsAsync(string username)
    {
        if (!Config.ClanFeatureEnabled)
        {
            return;
        }

        var member = Clan.ClanRosterStore.FindTrusted(username, Config.BattlenetServer);
        if (member is null || string.IsNullOrEmpty(member.Rank))
        {
            return;
        }

        var rank = Clan.ClanRankStore.Find(member.Rank);
        if (rank is null)
        {
            return;
        }

        if (rank.AutoBan)
        {
            var reason = string.IsNullOrWhiteSpace(rank.AutoBanMessage) ? "" : $" {rank.AutoBanMessage}";
            await SendChatCommandAsync($"/ban {username}{reason}").ConfigureAwait(false);
        }
        else if (rank.AutoKick)
        {
            var reason = string.IsNullOrWhiteSpace(rank.AutoKickMessage) ? "" : $" {rank.AutoKickMessage}";
            await SendChatCommandAsync($"/kick {username}{reason}").ConfigureAwait(false);
        }

        if (rank.HasAutoWhisper && ShouldSendAutoWhisper(member, rank.AutoWhisperFrequency))
        {
            await SendChatCommandAsync($"/w {username} {rank.AutoWhisperMessage}").ConfigureAwait(false);
            member.LastAutoWhisperUtc = DateTime.UtcNow;
            Clan.ClanRosterStore.Save();
        }
    }

    private static bool ShouldSendAutoWhisper(Clan.ClanMember member, Clan.AutoWhisperFrequency frequency)
    {
        if (member.LastAutoWhisperUtc is not { } last)
        {
            return true;
        }

        return frequency switch
        {
            Clan.AutoWhisperFrequency.EveryTime => true,
            Clan.AutoWhisperFrequency.Daily => DateTime.UtcNow - last >= TimeSpan.FromHours(24),
            Clan.AutoWhisperFrequency.Once => false,
            _ => false,
        };
    }

    private DateTime _lastChannelRecoveryAttemptUtc = DateTime.MinValue;

    /// <summary>
    /// "No one hears you." is PVPGN's reply to a chat send when the server
    /// no longer considers this connection to be in any channel at all —
    /// seen in practice when another client sharing the same channel (e.g.
    /// its effective operator, by naming convention like a channel named
    /// after that account) disconnects and the server silently drops
    /// remaining members without sending a Leave/rejoin event this bot would
    /// otherwise react to. Auto-rejoins the configured home channel to
    /// recover, cooled down to once per 10 seconds so a channel that's
    /// genuinely broken (not just desynced) can't turn into a rejoin-flood.
    /// </summary>
    private async Task RecoverFromChannelDesyncAsync()
    {
        var now = DateTime.UtcNow;
        if (now - _lastChannelRecoveryAttemptUtc < TimeSpan.FromSeconds(10))
        {
            return;
        }

        _lastChannelRecoveryAttemptUtc = now;
        LogInfo("Lost channel membership unexpectedly — rejoining home channel.");
        await JoinHomeAsync().ConfigureAwait(false);
    }

    private Task HandleNewsInfo(byte[] frame) => Task.CompletedTask;

    /// <summary>
    /// The server is asking for ExtraWork compliance — Blizzard's bot-
    /// detection mechanism (see BncsPacketId.SID_REQUIREDWORK's remarks).
    /// Deliberately not implemented: doing so means either running a
    /// server-provided native DLL (a real security risk on its own) or
    /// faking compliance, which is detection evasion against a real
    /// anti-bot system on a live commercial service — not something this
    /// project will do. Logged plainly rather than silently ignored, since
    /// on official Battle.net this is usually followed by the server
    /// dropping the connection some time later, and a silent, unexplained
    /// disconnect is worse than an honest one. PVPGN servers essentially
    /// never send this, since it's not part of core protocol compatibility.
    /// </summary>
    private Task HandleRequiredWork()
    {
        LogWarning(
            "Server requested ExtraWork compliance (Blizzard's anti-bot check) — not implemented by design. " +
            "This connection may be dropped by the server after a while; that's expected on official Battle.net " +
            "and isn't something this bot will work around.");
        return Task.CompletedTask;
    }

    private async Task HandleLegacyCreateAccountReplyAsync()
    {
        LogInfo("Account created! Reconnecting with your new account...");
        _bncs.Close();
        LogInfo($"Battle.net Login Server connecting to {Config.BnlsServer}...");
        await _bnls.ConnectAsync(Config.BnlsServer, Config.BnlsPort).ConfigureAwait(false);
    }

    private Task HandleSetEmail()
    {
        LogInfo("Please set an email address on this account from the game client; Invigoration does not register accounts.");
        return Task.CompletedTask;
    }

    private Task HandleFriendsList(byte[] frame)
    {
        _friends.Clear();
        _friends.AddRange(FriendsListParser.ParseFriendsList(frame));
        FriendsListUpdated?.Invoke(_friends);
        return Task.CompletedTask;
    }

    private Task HandleFriendsUpdate(byte[] frame)
    {
        var (entryNumber, update) = FriendsListParser.ParseFriendsUpdate(frame);
        if (entryNumber < _friends.Count)
        {
            var existing = _friends[entryNumber];
            _friends[entryNumber] = existing with
            {
                Status = update.Status,
                Location = update.Location,
                ProductCode = update.ProductCode,
                LocationName = update.LocationName,
            };
            FriendsListUpdated?.Invoke(_friends);
        }

        return Task.CompletedTask;
    }

    private Task HandleFriendsAdd(byte[] frame)
    {
        _friends.Add(FriendsListParser.ParseFriendsAdd(frame));
        FriendsListUpdated?.Invoke(_friends);
        return Task.CompletedTask;
    }

    private Task HandleFriendsRemove(byte[] frame)
    {
        var entryNumber = FriendsListParser.ParseFriendsRemove(frame);
        if (entryNumber < _friends.Count)
        {
            _friends.RemoveAt(entryNumber);
            FriendsListUpdated?.Invoke(_friends);
        }

        return Task.CompletedTask;
    }

    /// <summary>Moves the friend at the old position to the new one, shifting everything between — matches bnetdocs' description of this packet's effect exactly.</summary>
    private Task HandleFriendsPosition(byte[] frame)
    {
        var (oldEntry, newEntry) = FriendsListParser.ParseFriendsPosition(frame);
        if (oldEntry < _friends.Count && newEntry < _friends.Count)
        {
            var entry = _friends[oldEntry];
            _friends.RemoveAt(oldEntry);
            _friends.Insert(newEntry, entry);
            FriendsListUpdated?.Invoke(_friends);
        }

        return Task.CompletedTask;
    }
}
