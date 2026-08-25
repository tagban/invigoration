using Invigoration.Core.Chat;
using Invigoration.Core.Networking;
using Invigoration.Core.Protocol;

namespace Invigoration.Core;

/// <summary>
/// The classic BNCS 1.0 "profile" (SID_READUSERDATA/SID_WRITEUSERDATA, 0x26/0x27) — the
/// Sex/Age/Location/Description fields a real Battle.net chat client shows on right-click ->
/// Profile, editable for your own account. Field layout per bnetdocs.org, not yet confirmed
/// against a live server this session (unlike the rest of BotEngine.Bncs.cs, which has been) —
/// worth a real live test before relying on it.
/// </summary>
public sealed partial class BotEngine
{
    private static readonly string[] ProfileReadKeys = ["profile\\sex", "profile\\age", "profile\\location", "profile\\description"];

    private TaskCompletionSource<ProfileInfo>? _pendingProfileRequest;
    private string? _pendingProfileAccount;

    /// <summary>
    /// Requests another account's (or your own) profile fields. Only one request can be in
    /// flight at a time per bot — good enough for one Profile window at a time, which is the
    /// only caller. Returns null on timeout (e.g. the server doesn't support this for the
    /// current product, or simply never replies).
    /// </summary>
    public async Task<ProfileInfo?> RequestProfileAsync(string account, TimeSpan? timeout = null)
    {
        var tcs = new TaskCompletionSource<ProfileInfo>(TaskCreationOptions.RunContinuationsAsynchronously);
        _pendingProfileRequest = tcs;
        _pendingProfileAccount = account;

        var writer = new PacketWriter()
            .WriteDword(1) // number of accounts
            .WriteDword((uint)ProfileReadKeys.Length)
            .WriteDword(0) // request id, echoed back verbatim; unused since only one request is ever in flight
            .WriteNTString(account);
        foreach (var key in ProfileReadKeys)
        {
            writer.WriteNTString(key);
        }

        await SendBncsAsync(writer, BncsPacketId.SID_READUSERDATA).ConfigureAwait(false);

        var winner = await Task.WhenAny(tcs.Task, Task.Delay(timeout ?? TimeSpan.FromSeconds(10))).ConfigureAwait(false);
        if (winner != tcs.Task)
        {
            _pendingProfileRequest = null;
            return null;
        }

        return await tcs.Task.ConfigureAwait(false);
    }

    /// <summary>
    /// Writes your own profile (SID_WRITEUSERDATA has no reply — fire and forget). Age is
    /// deliberately excluded: classic clients only ever let you edit Sex/Location/Description,
    /// age comes from the account's birthdate on file with Blizzard and isn't writable this way.
    /// </summary>
    public Task WriteProfileAsync(string sex, string location, string description)
    {
        var writer = new PacketWriter()
            .WriteDword(1) // number of accounts
            .WriteDword(3) // number of keys
            .WriteNTString(OwnChatIdentity ?? Config.Username)
            .WriteNTString("profile\\sex")
            .WriteNTString("profile\\location")
            .WriteNTString("profile\\description")
            .WriteNTString(sex)
            .WriteNTString(location)
            .WriteNTString(description);
        return SendBncsAsync(writer, BncsPacketId.SID_WRITEUSERDATA);
    }

    private Task HandleReadUserData(byte[] frame)
    {
        var reader = BncsConnection.GetPayloadReader(frame);
        var numKeys = reader.ReadDword();
        var numAccounts = reader.ReadDword();
        reader.ReadDword(); // request id, echoed back; unused

        var values = new string[numKeys * numAccounts];
        for (var i = 0; i < values.Length; i++)
        {
            values[i] = reader.ReadNTString();
        }

        if (_pendingProfileRequest is { } tcs && values.Length >= ProfileReadKeys.Length)
        {
            tcs.TrySetResult(new ProfileInfo(_pendingProfileAccount ?? "", values[0], values[1], values[2], values[3]));
            _pendingProfileRequest = null;
        }

        return Task.CompletedTask;
    }
}
