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

    public ChatLineViewModel(IEnumerable<ChatLogSegment> segments, Bitmap? icon = null)
    {
        Segments = segments.Select(s => new ChatSegmentViewModel(s.Text, s.Color)).ToList();
        Icon = icon;
    }

    public ChatLineViewModel(string text, RgbColor color, Bitmap? icon = null)
    {
        Segments = [new ChatSegmentViewModel(text, color)];
        Icon = icon;
    }
}
