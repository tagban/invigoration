using Avalonia.Media;
using Avalonia.Media.Imaging;
using Invigoration.Core.Chat;

namespace Invigoration.App.Models;

public sealed class ChatSegmentViewModel(string text, RgbColor color)
{
    public string Text { get; } = text;

    public IBrush Brush { get; } = new SolidColorBrush(Color.FromRgb(color.R, color.G, color.B));
}

public sealed class ChatLineViewModel
{
    public IReadOnlyList<ChatSegmentViewModel> Segments { get; }

    /// <summary>The speaker's game/client icon, shown before the text — see BotConfig.ShowUserIconsInChat. Null on every line that isn't a Talk/Emote from a real user, or when the toggle is off.</summary>
    public Bitmap? Icon { get; }

    /// <summary>
    /// Show <see cref="Icon"/> as a large portrait with the speaker's name above the word-wrapped
    /// text, rather than a small inline icon (BotConfig.FullChatPortraits, SC2). The first segment
    /// is then the name.
    /// </summary>
    public bool LargePicture { get; }

    /// <summary>How many leading segments are the speaker's name (with its dimmed code and ": "), shown as the Full layout's heading.</summary>
    public int NameSegments { get; }

    /// <summary>The speaker's SC2 clan tag, without brackets, shown before their name in the Full layout; empty when none.</summary>
    public string ClanTag { get; init; } = "";

    public ChatLineViewModel(IEnumerable<ChatLogSegment> segments, Bitmap? icon = null, bool largePicture = false, int nameSegments = 1)
    {
        Segments = segments.Select(s => new ChatSegmentViewModel(s.Text, s.Color)).ToList();
        Icon = icon;
        NameSegments = Math.Clamp(nameSegments, 1, Math.Max(1, Segments.Count - 1));
        LargePicture = largePicture && icon is not null && Segments.Count > NameSegments;
    }

    public ChatLineViewModel(string text, RgbColor color, Bitmap? icon = null)
    {
        Segments = [new ChatSegmentViewModel(text, color)];
        Icon = icon;
    }
}
