namespace Invigoration.Core.Networking;

/// <summary>
/// Framing for Battle.net/PVPGN's older plain-text "Chat" connection type
/// (selected by sending byte 0x03 right after connecting, as opposed to
/// 0x01 for the normal binary Game/BNCS protocol) — a line-based, telnet-
/// style interface some PVPGN networks still run (e.g. war.pianka.io).
/// Frames on the first 0x0A or 0x0D byte found (bnetdocs notes either can
/// terminate a line); "\r\n"/"\n\r" pairs just produce one extra empty
/// frame, which callers skip rather than something this framing layer
/// needs to special-case.
/// </summary>
public sealed class ChatTelnetConnection : FramedTcpClient
{
    protected override int? TryGetFrameLength(IReadOnlyList<byte> buffer)
    {
        for (var i = 0; i < buffer.Count; i++)
        {
            if (buffer[i] is 0x0A or 0x0D)
            {
                return i + 1;
            }
        }

        return null;
    }

    /// <summary>
    /// Sends the connection-type/login-subtype selector bytes (0x03, 0x04) that open the
    /// handshake, before any line-based text is exchanged. They go out as ONE write on purpose:
    /// probed live against war.pianka.io (2026-09-15), sending them as two separate writes got the
    /// connection reset every time, while a single write was accepted every time — the server
    /// evidently reads both selector bytes in one go and treats a lone 0x03 segment as a malformed
    /// opening. The original version of this code sent them separately, which is the likeliest
    /// reason this connection type was written off as unreliable and removed back in August.
    /// </summary>
    public Task SendHandshakeAsync(CancellationToken cancellationToken = default) =>
        SendAsync([0x03, 0x04], cancellationToken);

    /// <summary>Sends one line of plain text, terminated with "\r\n".</summary>
    public Task SendLineAsync(string text, CancellationToken cancellationToken = default) =>
        SendAsync(System.Text.Encoding.UTF8.GetBytes(text + "\r\n"), cancellationToken);

    /// <summary>Decodes a received frame to text with its trailing line-terminator byte(s) stripped.</summary>
    public static string DecodeLine(byte[] frame) =>
        System.Text.Encoding.UTF8.GetString(frame).TrimEnd('\r', '\n');
}
