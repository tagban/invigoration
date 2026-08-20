using Invigoration.Core.Auth;
using Invigoration.Core.Networking;
using Invigoration.Core.Protocol;

namespace Invigoration.Core;

/// <summary>Port of modBNLS.bas's ParseBNLS. Frames are passed whole (header included) to keep offsets 1:1 with the original.</summary>
public sealed partial class BotEngine
{
    private Task HandleBnlsPacket(byte[] frame)
    {
        var id = (BnlsPacketId)BnlsConnection.GetPacketId(frame);
        LogDebug($"BNLS recv 0x{(byte)id:X2} ({id}), {frame.Length} bytes: {ToHexDump(frame)}");

        return id switch
        {
            BnlsPacketId.BNLS_VERSIONCHECKEX2 => HandleVersionCheckEx2Async(frame),
            BnlsPacketId.BNLS_VERSIONCHECK => HandleVersionCheckAsync(),
            BnlsPacketId.BNLS_CREATEACCOUNT => HandleCreateAccountReplyAsync(frame),
            BnlsPacketId.BNLS_LOGONCHALLENGE => HandleLogonChallengeReplyAsync(frame),
            BnlsPacketId.BNLS_LOGONPROOF => HandleLogonProofReplyAsync(frame),
            BnlsPacketId.BNLS_CDKEY_EX => HandleCdKeyExReplyAsync(frame),
            BnlsPacketId.BNLS_CDKEY => HandleCdKeyReplyAsync(frame),
            BnlsPacketId.BNLS_AUTHORIZE => HandleAuthorizeReplyAsync(frame),
            BnlsPacketId.BNLS_HASHDATA => HandleHashDataReplyAsync(frame),
            BnlsPacketId.BNLS_AUTHORIZEPROOF => HandleAuthorizeProofReplyAsync(),
            BnlsPacketId.BNLS_REQUESTVERSIONBYTE => HandleRequestVersionByteReplyAsync(frame),
            _ => Task.CompletedTask,
        };
    }

    private async Task HandleVersionCheckEx2Async(byte[] frame)
    {
        var reader = BnlsConnection.GetPayloadReader(frame);
        _ = reader.ReadBoolean(); // Success
        _auth.ExeVersion = reader.ReadDword();
        _auth.ExeChecksum = reader.ReadDword();
        _auth.ExeInfo = reader.ReadNTString();
        reader.ReadDword(); // Cookie, unused
        reader.ReadDword(); // Version code, unused

        if (!BncsProduct.RequiresCdKey(Config.Product))
        {
            await SendAuthCheckAsync().ConfigureAwait(false);
            return;
        }

        if (TryHashCdKeysLocally())
        {
            await SendAuthCheckAsync().ConfigureAwait(false);
            return;
        }

        // Fall back to BNLS for key formats CdKeyDecoder doesn't decode
        // locally (currently Warcraft III/TFT's 26-character key, which
        // already depends on BNLS for its NLS/SRP login regardless) or if a
        // key fails to decode for any other reason — BNLS/the server will
        // reject a genuinely bad key the same way either path.
        if (BncsProduct.RequiresExpansionCdKey(Config.Product))
        {
            // Confirmed against bnetdocs's BNLS_CDKEY_EX request layout:
            // Cookie(DWORD), KeyCount(BYTE), Flags(DWORD), then — since flags
            // here is CDKEY_SAME_SESSION_KEY (0x1) with GIVEN_SESSION_KEY
            // (0x2) unset — one shared server session key (DWORD) and no
            // client session keys, followed by one NTString per key. The
            // original VB6 (and this port, until now) only ever sent one
            // key string despite declaring KeyCount=2 — the actual bug
            // behind "there's no space for a secondary product key".
            var writer = new PacketWriter()
                .WriteDword(0) // Cookie
                .WriteByte(2) // Number of CD-keys
                .WriteDword(1) // Flags: CDKEY_SAME_SESSION_KEY
                .WriteDword(_auth.ServerToken) // shared server session key
                .WriteNTString(Config.CdKey)
                .WriteNTString(Config.ExpansionCdKey);
            await SendBnlsAsync(writer, BnlsPacketId.BNLS_CDKEY_EX).ConfigureAwait(false);
        }
        else
        {
            var writer = new PacketWriter().WriteDword(_auth.ServerToken).WriteNTString(Config.CdKey);
            await SendBnlsAsync(writer, BnlsPacketId.BNLS_CDKEY).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Computes the SID_AUTH_CHECK CD-key block(s) locally via CdKeyDecoder
    /// for the classic 13-digit and modern 16-character key formats, instead
    /// of asking BNLS_CDKEY/BNLS_CDKEY_EX to do it (avoiding sending the
    /// plaintext key over BNLS's unencrypted protocol). Generates a fresh
    /// client token ourselves — normally BNLS's job — since we're no longer
    /// asking it to. Returns false (leaving _auth untouched) if any required
    /// key doesn't decode, so the caller can fall back to BNLS.
    /// </summary>
    private bool TryHashCdKeysLocally()
    {
        var primary = CdKeyDecoder.Decode(Config.CdKey);
        if (primary is null)
        {
            return false;
        }

        DecodedCdKey? expansion = null;
        if (BncsProduct.RequiresExpansionCdKey(Config.Product))
        {
            expansion = CdKeyDecoder.Decode(Config.ExpansionCdKey);
            if (expansion is null)
            {
                return false;
            }
        }

        var clientToken = (uint)Random.Shared.Next();
        var blocks = new List<byte>(primary.Value.GetAuthCheckBlock(Config.CdKey.Trim().Length, clientToken, _auth.ServerToken));
        if (expansion is not null)
        {
            blocks.AddRange(expansion.Value.GetAuthCheckBlock(Config.ExpansionCdKey.Trim().Length, clientToken, _auth.ServerToken));
        }

        _auth.ClientToken = clientToken;
        _auth.CdKeyHash = blocks.ToArray();
        return true;
    }

    private Task HandleVersionCheckAsync()
    {
        var writer = new PacketWriter().WriteDword(_auth.ServerToken).WriteNTString(Config.CdKey);
        return SendBnlsAsync(writer, BnlsPacketId.BNLS_CDKEY);
    }

    private Task HandleCreateAccountReplyAsync(byte[] frame)
    {
        var reader = BnlsConnection.GetPayloadReader(frame);
        var hashedAccountData = reader.ReadRaw(reader.Remaining);
        var writer = new PacketWriter().WriteBytes(hashedAccountData).WriteNTString(Config.Username);
        return SendBncsAsync(writer, BncsPacketId.SID_AUTH_ACCOUNTCREATE);
    }

    private Task HandleLogonChallengeReplyAsync(byte[] frame)
    {
        var reader = BnlsConnection.GetPayloadReader(frame);
        var clientPublicKeyA = reader.ReadRaw(reader.Remaining);
        var writer = new PacketWriter().WriteBytes(clientPublicKeyA).WriteNTString(Config.Username);
        return SendBncsAsync(writer, BncsPacketId.SID_AUTH_ACCOUNTLOGON);
    }

    private Task HandleLogonProofReplyAsync(byte[] frame)
    {
        var reader = BnlsConnection.GetPayloadReader(frame);
        var proofM1 = reader.ReadRaw(reader.Remaining);
        var writer = new PacketWriter().WriteBytes(proofM1);
        return SendBncsAsync(writer, BncsPacketId.SID_AUTH_ACCOUNTLOGONPROOF);
    }

    /// <summary>
    /// Confirmed against bnetdocs's BNLS_CDKEY_EX reply layout: Cookie(DWORD),
    /// NumberRequested(BYTE), NumberSucceeded(BYTE), BitMask(DWORD), then per
    /// successful key: ClientSessionKey(DWORD) + CdKeyData(9 DWORDs = 36
    /// bytes). SID_AUTH_CHECK wants both keys' 36-byte blocks concatenated
    /// back-to-back. The previous single-key version of this handler (and
    /// its VB6-derived "big-endian token" quirk) was never exercised against
    /// a real expansion-product account — worth confirming live.
    /// </summary>
    private Task HandleCdKeyExReplyAsync(byte[] frame)
    {
        var reader = BnlsConnection.GetPayloadReader(frame);
        reader.Skip(4); // Cookie
        var numberRequested = reader.ReadByte();
        reader.Skip(1); // Number succeeded
        reader.Skip(4); // Bit mask

        var combinedHash = new List<byte>();
        for (var i = 0; i < numberRequested; i++)
        {
            var clientSessionKey = reader.ReadDword();
            if (i == 0)
            {
                _auth.ClientToken = clientSessionKey;
            }

            combinedHash.AddRange(reader.ReadRaw(36));
        }

        _auth.CdKeyHash = combinedHash.ToArray();
        return SendAuthCheckAsync();
    }

    private Task HandleCdKeyReplyAsync(byte[] frame)
    {
        var reader = BnlsConnection.GetPayloadReader(frame);
        if (!reader.ReadBoolean())
        {
            return Task.CompletedTask;
        }

        _auth.ClientToken = reader.ReadDword();
        _auth.CdKeyHash = reader.ReadRaw(36);
        return SendAuthCheckAsync();
    }

    private Task SendAuthCheckAsync()
    {
        var numKeys = !BncsProduct.RequiresCdKey(Config.Product) ? 0u
            : BncsProduct.RequiresExpansionCdKey(Config.Product) ? 2u : 1u;
        var writer = new PacketWriter()
            .WriteDword(_auth.ClientToken)
            .WriteDword(_auth.ExeVersion)
            .WriteDword(_auth.ExeChecksum)
            .WriteDword(numKeys)
            .WriteDword(0) // spawn
            .WriteBytes(_auth.CdKeyHash)
            .WriteNTString(_auth.ExeInfo)
            .WriteNTString(Config.Username);
        return SendBncsAsync(writer, BncsPacketId.SID_AUTH_CHECK);
    }

    private Task HandleAuthorizeReplyAsync(byte[] frame)
    {
        var reader = BnlsConnection.GetPayloadReader(frame);
        var challenge = reader.ReadDword();
        var response = Auth.BnlsChecksum.Compute(BnlsClientName, challenge);
        return SendBnlsAsync(new PacketWriter().WriteDword(response), BnlsPacketId.BNLS_AUTHORIZEPROOF);
    }

    private async Task HandleAuthorizeProofReplyAsync()
    {
        var productByte = BncsProduct.GetBnlsProductByte(Config.Product) ?? 0;
        await SendBnlsAsync(new PacketWriter().WriteDword(productByte), BnlsPacketId.BNLS_REQUESTVERSIONBYTE)
            .ConfigureAwait(false);

        if (Config.Product == BncsProduct.Warcraft3)
        {
            await SendBnlsAsync(new PacketWriter().WriteDword(2), BnlsPacketId.BNLS_CHOOSENLSREVISION)
                .ConfigureAwait(false);
        }
    }

    private async Task HandleRequestVersionByteReplyAsync(byte[] frame)
    {
        // Confirmed against multiple live captures with a valid product byte:
        // the reply payload is two DWORDs — the product byte echoed back,
        // then the actual version byte — matching modBNET.bas's Mid(Data,8,4)
        // offset (skip the first payload DWORD, read the second). An earlier
        // "fix" here mistakenly assumed a single-DWORD reply based on a
        // degenerate reply captured while Config.Product was corrupted
        // (product byte 0, which BNLS doesn't recognize and replies to with
        // a minimal 4-byte payload) — not representative of the real format.
        var reader = BnlsConnection.GetPayloadReader(frame);
        reader.Skip(4); // echoed product byte
        _auth.VersionByte = reader.ReadDword();

        _bncs.Close();
        LogInfo($"Battle.net connecting to {Config.BattlenetServer}...");
        await _bncs.ConnectAsync(Config.BattlenetServer, Config.BattlenetPort, proxy: BuildProxyOptions()).ConfigureAwait(false);
    }

    private async Task HandleHashDataReplyAsync(byte[] frame)
    {
        var reader = BnlsConnection.GetPayloadReader(frame);
        var hashResult = reader.ReadRaw(reader.Remaining);

        switch (_auth.HashPurpose)
        {
            case HashPurpose.AccountLogon:
            case HashPurpose.RealmLogon:
                await ContinueLogonHashFlowAsync(hashResult).ConfigureAwait(false);
                break;

            case HashPurpose.AccountCreate:
                await SendBncsAsync(
                    new PacketWriter().WriteBytes(hashResult).WriteNTString(Config.Username),
                    BncsPacketId.SID_CREATEACCOUNT).ConfigureAwait(false);
                break;

            case HashPurpose.ChangePassword:
                await ContinueChangePasswordHashFlowAsync(hashResult).ConfigureAwait(false);
                break;
        }
    }

    /// <summary>
    /// Old login system double-hash: stage 1 single-hashes the password (the
    /// reply we just got); we re-hash [ClientToken][ServerToken][singleHash]
    /// to get the double-hash BNCS actually wants, then send it as either
    /// SID_LOGONRESPONSE2 (account logon) or SID_LOGONREALMEX (realm logon).
    /// </summary>
    private async Task ContinueLogonHashFlowAsync(byte[] hashResult)
    {
        _auth.HashStage++;

        if (_auth.HashStage == 1)
        {
            var writer = new PacketWriter()
                .WriteDword(0x1C)
                .WriteDword(1)
                .WriteDword(_auth.ClientToken)
                .WriteDword(_auth.ServerToken)
                .WriteBytes(hashResult);
            await SendBnlsAsync(writer, BnlsPacketId.BNLS_HASHDATA).ConfigureAwait(false);
            return;
        }

        if (_auth.HashStage != 2)
        {
            return;
        }

        _auth.HashStage = 0;
        if (_auth.HashPurpose == HashPurpose.AccountLogon)
        {
            var writer = new PacketWriter()
                .WriteDword(_auth.ClientToken)
                .WriteDword(_auth.ServerToken)
                .WriteBytes(hashResult)
                .WriteNTString(Config.Username);
            await SendBncsAsync(writer, BncsPacketId.SID_LOGONRESPONSE2).ConfigureAwait(false);
        }
        else
        {
            var writer = new PacketWriter()
                .WriteDword(_auth.ClientToken)
                .WriteBytes(hashResult)
                .WriteNTString(Config.Realm);
            await SendBncsAsync(writer, BncsPacketId.SID_LOGONREALMEX).ConfigureAwait(false);
        }
    }

    private async Task ContinueChangePasswordHashFlowAsync(byte[] hashResult)
    {
        _auth.HashStage++;

        switch (_auth.HashStage)
        {
            case 1:
                var doubleHashRequest = new PacketWriter()
                    .WriteDword(0x1C)
                    .WriteDword(1)
                    .WriteDword(_auth.ClientToken)
                    .WriteDword(_auth.ServerToken)
                    .WriteBytes(hashResult);
                await SendBnlsAsync(doubleHashRequest, BnlsPacketId.BNLS_HASHDATA).ConfigureAwait(false);
                break;

            case 2:
                _auth.PendingOldPasswordDoubleHash = hashResult;
                await SendPasswordHashRequestAsync(_auth.NewPassword).ConfigureAwait(false);
                break;

            case 3:
                var writer = new PacketWriter()
                    .WriteDword(_auth.ClientToken)
                    .WriteDword(_auth.ServerToken)
                    .WriteBytes(_auth.PendingOldPasswordDoubleHash)
                    .WriteBytes(hashResult)
                    .WriteNTString(Config.Username);
                await SendBncsAsync(writer, BncsPacketId.SID_CHANGEPASSWORD).ConfigureAwait(false);
                _auth.HashStage = 0;
                break;
        }
    }
}
