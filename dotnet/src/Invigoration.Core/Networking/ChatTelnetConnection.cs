namespace Invigoration.Core.Networking;

/// <summary>
/// Framing for Battle.net/PVPGN's older plain-text "Chat" connection type
/// (selected by sending byte 0x03 right after connecting, as opposed to
/// 0x01 for the normal binary Game/BNCS protocol) — a line-based, telnet-
/// style interface some PVPGN networks still run (e.g. eurobattle.net).
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

    /// <summary>Sends a single raw byte — used for the two connection-type/login-subtype selector bytes (0x03, 0x04) at the very start of the handshake, before any line-based text is exchanged.</summary>
    public Task SendByteAsync(byte b, CancellationToken cancellationToken = default) =>
        SendAsync([b], cancellationToken);

    /// <summary>Sends one line of plain text, terminated with "\r\n".</summary>
    public Task SendLineAsync(string text, CancellationToken cancellationToken = default) =>
        SendAsync(System.Text.Encoding.UTF8.GetBytes(text + "\r\n"), cancellationToken);

    /// <summary>Decodes a received frame to text with its trailing line-terminator byte(s) stripped.</summary>
    public static string DecodeLine(byte[] frame) =>
        System.Text.Encoding.UTF8.GetString(frame).TrimEnd('\r', '\n');
}
