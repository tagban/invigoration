namespace Invigoration.Core.Chat;

/// <summary>
/// The joke text-transform toggles ("fudd"/"canada") plus prepend/postpend — extracted out of
/// BotEngine.cs's own ApplyTextEffects so the exact same behavior is available to any other chat
/// client (Hotline's HotlineSessionViewModel) without duplicating the transform logic. BotEngine's
/// own ApplyTextEffects now just delegates here; behavior is unchanged. Order matters: the
/// Fudd/Canada transforms apply first, then prepend/postpend wrap the already-transformed text —
/// a signature-style postpend shouldn't get its own R's turned into W's.
/// </summary>
public static class ChatTextEffects
{
    public static string Apply(string text, bool fuddMode, bool canadaMode, string prependText, string postpendText)
    {
        if (fuddMode)
        {
            text = text.Replace('r', 'w').Replace('R', 'W');
        }

        if (canadaMode)
        {
            text += ", eh?";
        }

        if (prependText.Length > 0)
        {
            text = prependText + " " + text;
        }

        if (postpendText.Length > 0)
        {
            text += " " + postpendText;
        }

        return text;
    }
}
