namespace Invigoration.Core.Chat;

/// <summary>
/// Parses inline color codes (a non-breaking-space marker (U+00A0, typed as
/// Alt+0160) followed by a single letter, e.g. "&#0160;r" for red) into
/// colored text runs. Port of modcolors.bas's fBotColors.
///
/// The original re-colored from each marker to the end of the *entire*
/// RichTextBox contents (a side effect of mutating a live, ever-growing
/// control's SelStart/SelColor rather than a deliberate design), which
/// doesn't translate to a segment-based renderer. Here a color code instead
/// scopes to the rest of the single message being parsed, which is the
/// same visible result for how the bot actually uses it (one color code per
/// message, e.g. the "colors" command's help text).
/// </summary>
public static class ChatColorFormatter
{
    private const char Marker = ' ';

    public static IReadOnlyList<ChatLogSegment> Parse(string text, RgbColor defaultColor, ChatPalette? palette = null)
    {
        palette ??= ChatPalette.Invigoration;
        var segments = new List<ChatLogSegment>();
        var currentColor = defaultColor;
        var runStart = 0;

        var i = 0;
        while (i < text.Length)
        {
            if (text[i] == Marker && i + 1 < text.Length)
            {
                if (runStart < i)
                {
                    segments.Add(new ChatLogSegment(currentColor, text[runStart..i]));
                }

                currentColor = GetColor(text[i + 1], palette) ?? currentColor;
                i += 2;
                runStart = i;
                continue;
            }

            i++;
        }

        if (runStart < text.Length)
        {
            segments.Add(new ChatLogSegment(currentColor, text[runStart..]));
        }

        return segments;
    }

    /// <summary>
    /// These inline codes are a fixed, scheme-independent palette (like a
    /// crayon box the operator can reach for in a message) rather than part
    /// of the bnubot-derived ChatPalette roles — bnubot's color schemes only
    /// define roles for this port's own system messages (errors, whispers,
    /// etc.), not for an arbitrary in-chat color-code feature.
    /// </summary>
    private static RgbColor? GetColor(char code, ChatPalette palette) => code switch
    {
        'r' => RgbColor.FromWin32Bgr(0xFF), // vbRed
        'w' => palette.White,
        'q' => RgbColor.FromWin32Bgr(0x808080), // vbGrey
        'g' => new RgbColor(0x00, 0xCE, 0x00),
        'y' => new RgbColor(0xFF, 0xFF, 0x00),
        'b' => new RgbColor(0x44, 0x40, 0x9C),
        'o' => new RgbColor(0xFF, 0xA5, 0x00),
        'c' => new RgbColor(0x00, 0xD0, 0xD0),
        'p' => new RgbColor(0x8D, 0x00, 0xCE),
        'l' => new RgbColor(0xCE, 0xCE, 0x51),
        'e' => RgbColor.FromWin32Bgr(0x659DA8), // D2Beige2
        'k' => new RgbColor(0xFF, 0x69, 0xB4),
        _ => null,
    };
}
