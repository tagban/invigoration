using System.Globalization;

namespace Invigoration.Core.Chat;

/// <summary>
/// Parses one line of Battle.net/PVPGN's plain-text "Chat" connection type
/// (see <see cref="Networking.ChatTelnetConnection"/>) into the same
/// <see cref="ChatEvent"/> record the binary BNCS SID_CHATEVENT parser
/// produces, so BotEngine's whole downstream chat-event pipeline (roster
/// tracking, rank behaviors, trivia, command dispatch) is shared between
/// both connection types without duplicating it.
///
/// Event line format, confirmed against a live capture (2026-08):
///   "&lt;eventId&gt; USER &lt;username&gt; &lt;flags-hex&gt; &lt;statstring-or-tag&gt;"
///   "&lt;eventId&gt; TALK &lt;username&gt; &lt;flags-hex&gt; &quot;&lt;message&gt;&quot;"
///   "&lt;eventId&gt; CHANNEL &quot;&lt;channel name&gt;&quot;"
/// The numeric event IDs are exactly 1000 + the matching <see cref="ChatEventType"/>
/// wire value (e.g. 1001 = 1000 + ShowUser's 0x1, 1005 = 1000 + Talk's 0x5,
/// 1007 = 1000 + Channel's 0x7) — confirmed for USER/JOIN/TALK/CHANNEL by the
/// live capture; the rest (LEAVE/WHISPER/BROADCAST/USERFLAGS/WHISPERSENT/
/// INFO/ERROR/EMOTE below) follow the same pattern but are NOT independently
/// confirmed — bnetdocs' own page for this protocol says it's incomplete
/// ("For a full list of packets, contact me"), so treat those as
/// best-effort until verified live against a real Chat-protocol server.
/// </summary>
public static class ChatTelnetEventParser
{
    /// <summary>The one-time post-login confirmation line ("2010 NAME &lt;username&gt;") — not a ChatEvent, handled separately during the connect handshake.</summary>
    public const int NameConfirmationEventId = 2010;

    private static readonly Dictionary<int, ChatEventType> EventIdsByType = new()
    {
        [1001] = ChatEventType.ShowUser,
        [1002] = ChatEventType.Join,
        [1003] = ChatEventType.Leave,
        [1004] = ChatEventType.Whisper,
        [1005] = ChatEventType.Talk,
        [1006] = ChatEventType.Broadcast,
        [1007] = ChatEventType.Channel,
        [1009] = ChatEventType.UserFlags,
        [1010] = ChatEventType.WhisperSent,
        [1018] = ChatEventType.Info,
        [1019] = ChatEventType.Error,
        [1023] = ChatEventType.Emote,
    };

    /// <summary>Null for a line that isn't a recognized numbered event (login prompts, the 2010 NAME line, blank lines, or an event ID not in the table above).</summary>
    public static ChatEvent? TryParse(string line)
    {
        line = line.Trim();
        if (line.Length == 0)
        {
            return null;
        }

        var firstSpace = line.IndexOf(' ');
        if (firstSpace < 0 || !int.TryParse(line[..firstSpace], out var eventId))
        {
            return null;
        }

        if (!EventIdsByType.TryGetValue(eventId, out var type))
        {
            return null;
        }

        var afterEventId = line[(firstSpace + 1)..].TrimStart();
        var keywordSpace = afterEventId.IndexOf(' ');
        var fields = keywordSpace < 0 ? "" : afterEventId[(keywordSpace + 1)..].TrimStart();

        return type is ChatEventType.Channel or ChatEventType.Info or ChatEventType.Error
            ? new ChatEvent(type, "", 0, 0, Unquote(fields))
            : ParseUserEvent(type, fields);
    }

    /// <summary>"&lt;username&gt; &lt;flags-hex&gt; &lt;text, optionally quoted&gt;".</summary>
    private static ChatEvent? ParseUserEvent(ChatEventType type, string fields)
    {
        var parts = fields.Split(' ', 3);
        if (parts.Length < 2)
        {
            return null;
        }

        var flags = uint.TryParse(parts[1], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var parsedFlags)
            ? parsedFlags
            : 0;
        var text = parts.Length > 2 ? Unquote(parts[2]) : "";
        return new ChatEvent(type, parts[0], flags, 0, text);
    }

    private static string Unquote(string s)
    {
        s = s.Trim();
        return s.Length >= 2 && s[0] == '"' && s[^1] == '"' ? s[1..^1] : s;
    }
}
