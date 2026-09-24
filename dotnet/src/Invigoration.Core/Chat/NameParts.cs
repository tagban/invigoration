using System.Text.RegularExpressions;

namespace Invigoration.Core.Chat;

/// <summary>Battle.net 2.0 names end in a "#1234" code (BattleTags, SC2 characters); shown dimmed, it reads more like a chat.</summary>
public static partial class NameParts
{
    /// <summary>"Elesh#700" → ("Elesh", "#700"); a name with no code → (name, "").</summary>
    public static (string Name, string Code) Split(string name) =>
        CodeSuffix().Match(name) is { Success: true } match ? (match.Groups[1].Value, match.Groups[2].Value) : (name, "");

    [GeneratedRegex(@"^(.+?)(#\d+)$")]
    private static partial Regex CodeSuffix();
}
