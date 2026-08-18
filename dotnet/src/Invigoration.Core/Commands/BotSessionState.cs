namespace Invigoration.Core.Commands;

/// <summary>
/// Runtime toggles and counters the chat commands read/mutate. Replaces the
/// Public/Global variables scattered across modCommands.bas and globals.bas
/// (Canada, leetspeak, fudd, moo, debugmode, acceptinvites, idleMessage,
/// BanCount/KickCount/JoinCount, LastW/LastM/LastSW/LastSM, beforetext,
/// postpend, targetuser) — kept as one instance per <see cref="Invigoration.Core.BotEngine"/>.
/// </summary>
public sealed class BotSessionState
{
    public bool CanadaMode { get; set; }
    public bool LeetSpeakMode { get; set; }
    public bool FuddMode { get; set; }
    public bool MooMode { get; set; }
    public bool DebugMode { get; set; }
    public bool AcceptClanInvites { get; set; }

    public string IdleMessage { get; set; } = "";
    public int IdleTimeSetMinutes { get; set; }

    public string PrependText { get; set; } = "";
    public string PostpendText { get; set; } = "";

    /// <summary>Whisper-focus target set by the "user" command; empty means channel/no focus.</summary>
    public string TargetUser { get; set; } = "";

    public int BanCount { get; set; }
    public int KickCount { get; set; }
    public int JoinCount { get; set; }

    public string LastWhisperFromUser { get; set; } = "";
    public string LastWhisperFromText { get; set; } = "";
    public string LastWhisperSentUser { get; set; } = "";
    public string LastWhisperSentText { get; set; } = "";
}
