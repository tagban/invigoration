using Avalonia.Media;
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

    public ChatLineViewModel(IEnumerable<ChatLogSegment> segments)
    {
        Segments = segments.Select(s => new ChatSegmentViewModel(s.Text, s.Color)).ToList();
    }

    public ChatLineViewModel(string text, RgbColor color)
    {
        Segments = [new ChatSegmentViewModel(text, color)];
    }
}
