namespace Invigoration.Core.Discord;

/// <summary>
/// The shape an Invigoration Discord relay posts into Battle.net: <c>[Discord] name: message</c>.
/// One definition for both ends — the relaying bot writes it with <see cref="Format"/>, and every
/// Invigoration bot in the channel reads it back with <see cref="TryParse"/> to show the Discord
/// user's own name with a Discord logo instead of the relay account's name.
/// </summary>
public static class DiscordRelayLine
{
    public const string Prefix = "[Discord] ";

    /// <summary>Discord usernames are at most 32 characters; a little headroom for older "name#1234" forms.</summary>
    private const int MaxNameLength = 40;

    /// <summary>How a Discord user is named on the Battle.net side: "[Discord] name".</summary>
    public static string SpeakerName(string discordUser) => Prefix + discordUser;

    public static string Format(string discordUser, string message) => $"{SpeakerName(discordUser)}: {message}";

    /// <summary>Recognizes a relayed line. False for anything else, including a line that merely starts with the prefix but has no name, no ": " separator, or an implausibly long name.</summary>
    public static bool TryParse(string text, out string discordUser, out string message)
    {
        discordUser = message = "";
        if (!text.StartsWith(Prefix, StringComparison.Ordinal))
        {
            return false;
        }

        var separator = text.IndexOf(": ", Prefix.Length, StringComparison.Ordinal);
        if (separator < 0)
        {
            return false;
        }

        var name = text[Prefix.Length..separator];
        if (name.Trim().Length == 0 || name.Length > MaxNameLength)
        {
            return false;
        }

        discordUser = name;
        message = text[(separator + 2)..];
        return true;
    }

    /// <summary>The Discord name inside a "[Discord] name" speaker, or null for any other speaker.</summary>
    public static string? DiscordUserFromSpeaker(string speaker) =>
        speaker.StartsWith(Prefix, StringComparison.Ordinal) && speaker.Length > Prefix.Length ? speaker[Prefix.Length..] : null;
}
