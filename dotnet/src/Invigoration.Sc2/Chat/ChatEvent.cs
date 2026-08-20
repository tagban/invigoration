namespace Invigoration.Sc2.Chat;

/// <summary>
/// Events surfaced from an established chat session. Mirrors
/// core/src/chat/session.rs's ChatEvent — currently covers the variants
/// confirmed from the reference example client (core/examples/sc2-tui-bot.rs);
/// the remaining ones (Friends, BlockedAccounts, Activity, GroupInvitation,
/// PartyInvitation, GroupSummary, GroupSearch, PublicChannelCatalog,
/// ConferenceDirectory) aren't modeled yet since their exact field shapes
/// weren't confirmed against primary source at the time this was written —
/// add them once native/decode.rs's hand-written decoders are ported.
/// </summary>
public abstract record ChatEvent
{
    private ChatEvent()
    {
    }

    public sealed record Joined(byte ChannelIndex, ChatChannel Channel) : ChatEvent;

    public sealed record JoinRejected(ChatChannel Channel, string? Reason) : ChatEvent;

    public sealed record Roster(RosterSnapshot Snapshot) : ChatEvent;

    public sealed record MemberJoined(byte ChannelIndex, ChatUser User) : ChatEvent;

    public sealed record MemberLeft(byte ChannelIndex, ChatUser User) : ChatEvent;

    public sealed record Removed(byte ChannelIndex) : ChatEvent;

    public sealed record Message(byte ChannelIndex, ChatUser Sender, string Body) : ChatEvent;

    public sealed record Whisper(string Peer, string Body, bool Outgoing) : ChatEvent;

    public sealed record WhisperFailed(string Peer, string Reason) : ChatEvent;
}
