namespace Invigoration.Sc2.Chat;

/// <summary>
/// Which channel to join/address. Mirrors core/src/chat/session.rs's
/// ChatChannel enum. Public channels are identified by a numeric catalog id
/// (General's is a placeholder that gets remapped from the real public
/// channel catalog once it loads — see <see cref="DefaultPublicChannel"/>).
/// </summary>
public abstract record ChatChannel
{
    /// <summary>Placeholder id for "General" — the real numeric id is resolved from the public-channel catalog after connecting.</summary>
    public const ushort DefaultPublicChannel = 1028;

    private ChatChannel()
    {
    }

    public sealed record Public(ushort ChannelId) : ChatChannel;

    public sealed record Private(string Name) : ChatChannel;

    public sealed record Club(uint ClubId) : ChatChannel;

    public sealed record Party : ChatChannel;

    public static ChatChannel DefaultPublic() => new Public(DefaultPublicChannel);

    /// <summary>Human-readable title for a channel, matching core/src/chat/session.rs's channel_title.</summary>
    public string Title() => this switch
    {
        Public { ChannelId: DefaultPublicChannel } => "General",
        Public p => $"Public {p.ChannelId}",
        Private p => p.Name,
        Club c => $"Group {c.ClubId}",
        Party => "Party",
        _ => "Unknown",
    };
}
