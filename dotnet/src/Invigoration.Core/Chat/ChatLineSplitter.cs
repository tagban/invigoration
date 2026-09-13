namespace Invigoration.Core.Chat;

/// <summary>
/// Breaks outgoing classic Battle.net chat into lines the server will accept. SID_CHATCOMMAND
/// carries at most 223 characters of text plus its null terminator: official Battle.net trims
/// anything longer, but BNETDocs' Atlas treats a longer packet as a protocol violation and closes
/// the connection outright — a long trivia leaderboard or Discord relay was enough to knock a bot
/// offline there.
/// </summary>
/// <remarks>
/// Text is sent Latin-1 (one byte per char — see PacketWriter.WriteAscii), so characters and
/// bytes line up. Plain chat and the message-carrying commands (whisper and emote) are wrapped at
/// word boundaries with their command prefix repeated on every line, so a long whisper stays a
/// whisper to the same person. Any other slash command can't be meaningfully split and is cut to
/// the limit instead.
/// </remarks>
public static class ChatLineSplitter
{
    public const int MaxLineLength = 223;

    private static readonly string[] WhisperCommands = ["/w ", "/whisper ", "/m ", "/msg "];
    private static readonly string[] EmoteCommands = ["/me ", "/emote "];

    public static IReadOnlyList<string> Split(string text, int maxLength = MaxLineLength)
    {
        if (string.IsNullOrEmpty(text))
        {
            return [];
        }

        if (text.Length <= maxLength)
        {
            return [text];
        }

        var prefix = MessagePrefix(text);
        if (prefix is null)
        {
            return [text[..maxLength]];
        }

        var body = text[prefix.Length..];
        var room = maxLength - prefix.Length;
        if (room < 1)
        {
            return [text[..maxLength]];
        }

        return WrapWords(body, room).Select(line => prefix + line).ToList();
    }

    /// <summary>What has to be repeated on each wrapped line: "" for plain chat, "/w name " for a whisper, "/me " for an emote, or null for any other command (not splittable).</summary>
    private static string? MessagePrefix(string text)
    {
        if (text[0] != '/')
        {
            return "";
        }

        foreach (var command in EmoteCommands)
        {
            if (text.StartsWith(command, StringComparison.OrdinalIgnoreCase))
            {
                return text[..command.Length];
            }
        }

        foreach (var command in WhisperCommands)
        {
            if (!text.StartsWith(command, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var targetEnd = text.IndexOf(' ', command.Length);
            return targetEnd < 0 ? null : text[..(targetEnd + 1)];
        }

        return null;
    }

    private static IEnumerable<string> WrapWords(string body, int room)
    {
        var start = 0;
        while (start < body.Length)
        {
            if (body.Length - start <= room)
            {
                yield return body[start..];
                yield break;
            }

            // Break at the last space that fits; a single word longer than a whole line is cut.
            var breakAt = body.LastIndexOf(' ', start + room, room + 1);
            if (breakAt <= start)
            {
                yield return body.Substring(start, room);
                start += room;
                continue;
            }

            yield return body[start..breakAt];
            start = breakAt + 1;
            while (start < body.Length && body[start] == ' ')
            {
                start++;
            }
        }
    }
}
