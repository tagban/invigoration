using Invigoration.Sc2.Chat;

namespace Invigoration.Sc2.Connection;

/// <summary>Commands sent to a running client actor. Mirrors core/src/connection.rs's ClientCommand.</summary>
public abstract record ClientCommand
{
    private ClientCommand()
    {
    }

    public sealed record Connect(bool ForceInteractive, IReadOnlyList<ChatChannel> Channels) : ClientCommand;

    public sealed record Disconnect : ClientCommand;

    public sealed record SignOut : ClientCommand;

    public sealed record JoinChannel(ChatChannel Channel) : ClientCommand;

    public sealed record LeaveChannel(byte ChannelIndex) : ClientCommand;

    public sealed record SendMessage(byte ChannelIndex, string Body) : ClientCommand;

    public sealed record SendWhisper(WhisperTarget Target, string DisplayName, string Body) : ClientCommand;

    public sealed record AnswerGroupInvitation(uint ClubId, bool Accept) : ClientCommand;

    public sealed record AnswerPartyInvitation(byte ChannelIndex, bool Accept) : ClientCommand;

    public sealed record SearchGroups(string Query) : ClientCommand;

    public sealed record Quit : ClientCommand;
}
