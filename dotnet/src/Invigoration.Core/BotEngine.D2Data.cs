using Invigoration.Core.Config;
using Invigoration.Core.Networking;
using Invigoration.Core.Protocol;

namespace Invigoration.Core;

/// <summary>
/// Keeping the opted-in Diablo II data (D2EquipmentStore) current. After logging on to a trusted
/// server (D2EquipmentStore.TrustedServers — only us.bnet.cc for now), the bot asks SID_GETFILETIME
/// for each kept file; any the server has a newer copy of — or has for the first time, like an art
/// pack that wasn't ready when the user opted in — is fetched from that same server over BNFTP in
/// the background. Nothing is asked before the user opts in, and no other server is ever asked.
/// </summary>
public sealed partial class BotEngine
{
    /// <summary>Request ids for SID_GETFILETIME, one per kept file (the index into D2EquipmentStore.TrackedFiles), tagged so another reply can't be mistaken for one.</summary>
    private const uint D2FileTimeRequestBase = 0x44320000;

    /// <summary>Only a trusted server is ever asked, and only once the user has opted in.</summary>
    internal static bool ShouldCheckD2FileTimes(string server, bool optedIn) =>
        optedIn && D2EquipmentStore.IsTrustedServer(server);

    private async Task RequestD2FileTimesAsync()
    {
        if (!ShouldCheckD2FileTimes(Config.BattlenetServer, D2EquipmentStore.OptedIn))
        {
            return;
        }

        for (var i = 0; i < D2EquipmentStore.TrackedFiles.Count; i++)
        {
            await SendBncsAsync(
                new PacketWriter().WriteDword(D2FileTimeRequestBase + (uint)i).WriteDword(0).WriteNTString(D2EquipmentStore.TrackedFiles[i]),
                BncsPacketId.SID_GETFILETIME).ConfigureAwait(false);
        }
    }

    /// <summary>SID_GETFILETIME reply: (DWORD) request id, (DWORD) unknown, (FILETIME) file time, (STRING) file name. A time of 0 means the server doesn't have it.</summary>
    private Task HandleFileTimeReply(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        var requestId = reader.ReadDword();
        reader.ReadDword();
        var time = reader.ReadFileTime();
        var fileName = reader.ReadNTString();
        var fileTime = (long)(((ulong)time.High << 32) | time.Low);

        if (requestId - D2FileTimeRequestBase >= (uint)D2EquipmentStore.TrackedFiles.Count ||
            !D2EquipmentStore.IsTrustedServer(Config.BattlenetServer) ||
            !D2EquipmentStore.IsNewerOnServer(fileName, fileTime))
        {
            return Task.CompletedTask;
        }

        var host = Config.BattlenetServer;
        var port = Config.BattlenetPort > 0 ? Config.BattlenetPort : BnftpClient.DefaultPort;
        SafeFireAndForget(RefreshD2FileAsync(host, port, fileName), $"updating {fileName}");
        return Task.CompletedTask;
    }

    private async Task RefreshD2FileAsync(string host, int port, string fileName)
    {
        var (result, detail) = await D2EquipmentStore.RefreshFileAsync(host, port, fileName).ConfigureAwait(false);
        if (result == D2EquipmentDownloadResult.Saved)
        {
            LogInfo($"Updated Diablo II data: {fileName} ({detail}).");
        }
        else
        {
            LogDebug($"Diablo II data update for {fileName} didn't complete: {detail}.");
        }
    }
}
