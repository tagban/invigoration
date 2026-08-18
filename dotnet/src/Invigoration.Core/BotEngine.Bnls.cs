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

        if (BncsProduct.RequiresExpansionCdKey(Config.Product))
        {
            // NOTE: ported as-is from the original — a real dual-key
            // BNLS_CDKEY_EX request should carry two hashed keys, but only
            // one CD-key field exists here (and in the VB6 config form this
            // came from), matching a limitation already present upstream.
            var writer = new PacketWriter()
                .WriteDword(0)
                .WriteByte(2)
                .WriteDword(1)
                .WriteDword(_auth.ServerToken)
                .WriteNTString(Config.CdKey);
            await SendBnlsAsync(writer, BnlsPacketId.BNLS_CDKEY_EX).ConfigureAwait(false);
        }
        else
        {
            var writer = new PacketWriter().WriteDword(_auth.ServerToken).WriteNTString(Config.CdKey);
            await SendBnlsAsync(writer, BnlsPacketId.BNLS_CDKEY).ConfigureAwait(false);
        }
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

    private Task HandleCdKeyExReplyAsync(byte[] frame)
    {
        var reader = BnlsConnection.GetPayloadReader(frame);
        reader.Skip(10); // matches modBNLS.bas's derivation for this reply layout
        var clientTokenBytes = reader.ReadRaw(4);
        Array.Reverse(clientTokenBytes); // this reply's token is big-endian, unlike the rest of the protocol
        _auth.ClientToken = BitConverter.ToUInt32(clientTokenBytes);
        _auth.CdKeyHash = reader.ReadRaw(36);
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
        var numKeys = BncsProduct.RequiresExpansionCdKey(Config.Product) ? 2u : 1u;
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
        await _bncs.ConnectAsync(Config.BattlenetServer, Config.BattlenetPort).ConfigureAwait(false);
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
