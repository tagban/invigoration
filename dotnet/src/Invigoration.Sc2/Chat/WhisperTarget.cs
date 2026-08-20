namespace Invigoration.Sc2.Chat;

/// <summary>Who to address a whisper to. Mirrors the WhisperTarget cases used by core/src/native/protocol.rs's chat_whisper builder.</summary>
public abstract record WhisperTarget
{
    private WhisperTarget()
    {
    }

    public sealed record Presence(uint PresenceId) : WhisperTarget;

    public sealed record ToonName(string Name, byte Region, uint ProgramId, uint Realm) : WhisperTarget;

    public sealed record Account(uint AccountId) : WhisperTarget;

    public sealed record ToonHandle(uint ProgramId, byte Region, uint Realm, ulong Id) : WhisperTarget;
}
