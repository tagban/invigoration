using Avalonia.Media;
using Avalonia.Media.Imaging;

namespace Invigoration.App.ViewModels;

/// <summary>
/// One line in a session's chat log. Chat/status lines are plain (Username null); a real chat
/// message gets its sender split out and colored by rank — see
/// HotlineSessionViewModel.TryColorizeChatLine's remarks for why this has to be done by matching
/// against the live user list rather than reading a structured field (the protocol doesn't send
/// one; the server pre-formats the whole line as one string). UsernameIcon is set for a
/// Discord-relay message (see HotlineSessionViewModel.TryAppendDiscordRelayMessage) — the small
/// Discord mark shown before the (real, relayed) sender's name.
/// </summary>
public sealed record HotlineChatLine(string? Username, IBrush? UsernameColor, string Text, Bitmap? UsernameIcon = null)
{
    public static HotlineChatLine Plain(string text) => new(null, null, text);

    /// <summary>Reconstructs the original full line — used by "Copy Log" so a colorized message still copies as plain, readable text.</summary>
    public string FullText => Username is null ? Text : Username + Text;
}
