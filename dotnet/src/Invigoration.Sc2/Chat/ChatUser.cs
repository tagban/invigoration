namespace Invigoration.Sc2.Chat;

/// <summary>Presence state for a channel member, per core/src/chat/session.rs.</summary>
public enum PresenceState
{
    Unknown,
    Offline,
    Online,
    Away,
    Busy,
}

/// <summary>One member of a chat channel. Mirrors core/src/chat/session.rs's ChatUser.</summary>
public sealed record ChatUser(
    uint Handle,
    uint? PresenceId,
    string? Name,
    string? ClanTag,
    PresenceState Presence)
{
    /// <summary>The name shown in the UI: clan tag prefix (if any) plus the name with any trailing "#1234" BattleTag discriminator stripped.</summary>
    public string VisibleName()
    {
        var name = StripCharacterCode(Name) ?? $"Player {Handle}";
        return ClanTag is { Length: > 0 } tag ? $"<{tag}>{name}" : name;
    }

    private static string? StripCharacterCode(string? name)
    {
        if (name is null)
        {
            return null;
        }

        var hashIndex = name.IndexOf('#');
        return hashIndex < 0 ? name : name[..hashIndex];
    }
}

/// <summary>The full member list for one joined channel at a point in time.</summary>
public sealed record RosterSnapshot(byte ChannelIndex, bool InitialComplete, IReadOnlyList<ChatUser> Users);
