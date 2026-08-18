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
            BncsPacketId.SID_CHATEVENT => HandleChatEvent(frame),
            BncsPacketId.SID_NEWS_INFO => HandleNewsInfo(frame),
            BncsPacketId.SID_CREATEACCOUNT => HandleLegacyCreateAccountReplyAsync(),
            BncsPacketId.SID_SETEMAIL => HandleSetEmail(),
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
        if (reader.ReadWord() == 0)
        {
            LogInfo("Account created! Connecting with your new account...");
            var writer = new PacketWriter().WriteNTString(Config.Username).WriteNTString(Config.Password);
            await SendBnlsAsync(writer, BnlsPacketId.BNLS_LOGONCHALLENGE).ConfigureAwait(false);
        }
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
        await _realm.ConnectAsync(host, RealmPort).ConfigureAwait(false);
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

    private async Task HandleChatEvent(byte[] frame)
    {
        var chatEvent = ChatEventParser.Parse(frame);
        ChatMessage?.Invoke(chatEvent);

        if (chatEvent.Type is ChatEventType.Talk or ChatEventType.Whisper)
        {
            if (chatEvent.Type == ChatEventType.Whisper)
            {
                _session.LastWhisperFromUser = chatEvent.Username;
                _session.LastWhisperFromText = chatEvent.Text;
            }

            await HandleCommandAsync(chatEvent.Username, chatEvent.Text).ConfigureAwait(false);
        }
    }

    private Task HandleNewsInfo(byte[] frame) => Task.CompletedTask;

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
}
