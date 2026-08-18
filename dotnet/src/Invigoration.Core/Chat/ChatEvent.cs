namespace Invigoration.Core.Chat;

/// <summary>SID_CHATEVENT event IDs (EID_*), per bnetdocs. Port of the ID_* constants in bnetbot.cls.</summary>
public enum ChatEventType : uint
{
    ShowUser = 0x1,
    Join = 0x2,
    Leave = 0x3,
    Whisper = 0x4,
    Talk = 0x5,
    Broadcast = 0x6,
    Channel = 0x7,
    UserFlags = 0x9,
    WhisperSent = 0xA,
    Info = 0x12,
    Error = 0x13,
    Emote = 0x17,
}

/// <summary>A parsed SID_CHATEVENT. Replaces bnetbot.cls's per-event-type VB6 events with one typed record.</summary>
public sealed record ChatEvent(ChatEventType Type, string Username, uint Flags, int Ping, string Text);

[Flags]
public enum UserFlags : uint
{
    None = 0,
    Blizzard = 0x1,
    Operator = 0x2,
    Speaker = 0x4,
    Admin = 0x8,
    NoUdp = 0x10,
    Squelched = 0x20,
    Special = 0x40, // "glasses" icon in the original
    GameFounder = 0x200000,
    InvigorationTeam = 0x80000,
    BnuBot = 0x800000,
    Hacker = 0x8000000,
    Warcraft3 = 0x80000000,
}
