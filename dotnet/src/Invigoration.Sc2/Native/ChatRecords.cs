namespace Invigoration.Sc2.Native;

/// <summary>Decoded payload of a ChatInviteNotify record (command 4). Mirrors native::model::ChatInvite.</summary>
public sealed record ChatInviteRecord(byte ChannelType, uint InviterPresence, byte ChannelIndex);

/// <summary>
/// Decoded payload of a JoinNotify2 record (command 27). Mirrors
/// native::model::ChatJoin. On failure only <see cref="Reason"/> (and
/// optionally <see cref="ChannelType"/>/<see cref="Token"/>) are populated;
/// on success everything except <see cref="Reason"/> may be.
/// </summary>
public sealed record ChatJoinRecord(
    bool Success,
    byte? ChannelIndex,
    uint? MemberHandle,
    byte? ChannelType,
    ushort? ChannelNameId,
    ushort? ChannelShardIndex,
    ushort? Reason,
    uint? Token);

/// <summary>Decoded payload of a MessageRecv record (command 11). Mirrors native::model::ChatMessage.</summary>
public sealed record ChatMessageRecord(uint MemberHandle, string Body, byte ChannelIndex);

/// <summary>Decoded payload of a WhisperRecv/WhisperEchoRecv record (commands 19/30). Mirrors native::model::ChatWhisper.</summary>
public sealed record ChatWhisperRecord(byte PeerRegion, uint PeerProgramId, uint PeerRealm, string PeerName, string Body);
