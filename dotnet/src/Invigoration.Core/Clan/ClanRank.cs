namespace Invigoration.Core.Clan;

public enum AutoWhisperFrequency
{
    /// <summary>Whisper every single time a member with this rank is seen.</summary>
    EveryTime,

    /// <summary>Whisper at most once per rolling 24 hours.</summary>
    Daily,

    /// <summary>Whisper only the very first time this member is ever seen holding this rank.</summary>
    Once,
}

/// <summary>
/// A predefined rank in the shared clan roster. Beyond labeling a member
/// (and being something PermissionLevel.Ranks can grant commands to), a
/// rank can carry automated behaviors the bot applies whenever it sees a
/// member holding it: a welcome whisper, or — for flagging troublemakers —
/// an automatic kick or ban. All behaviors are off by default; a rank with
/// none of them set behaves exactly like the old free-text rank string did.
/// </summary>
public sealed class ClanRank
{
    public string Name { get; set; } = "";

    /// <summary>Sent as a whisper whenever a member holding this rank is seen (channel join or the bot's own initial roster) — empty disables it.</summary>
    public string AutoWhisperMessage { get; set; } = "";

    public AutoWhisperFrequency AutoWhisperFrequency { get; set; } = AutoWhisperFrequency.Once;

    /// <summary>Automatically "/kick" a member holding this rank whenever they're seen — for flagging troublemakers without needing to watch for them manually.</summary>
    public bool AutoKick { get; set; }

    /// <summary>Optional reason appended to the auto-kick ("/kick username reason") — blank sends no reason.</summary>
    public string AutoKickMessage { get; set; } = "";

    /// <summary>Automatically "/ban" a member holding this rank whenever they're seen.</summary>
    public bool AutoBan { get; set; }

    /// <summary>Optional reason appended to the auto-ban ("/ban username reason") — blank sends no reason.</summary>
    public string AutoBanMessage { get; set; } = "";

    /// <summary>Canonical command names (see Commands.CommandCatalog) anyone holding this rank may use — replaces the old per-bot PermissionLevel system, so access now lives entirely on the rank instead of a separate config-only grant list. The bot master always has full access regardless of this.</summary>
    public List<string> AllowedCommands { get; set; } = [];

    public bool HasAutoWhisper => !string.IsNullOrWhiteSpace(AutoWhisperMessage);
}
