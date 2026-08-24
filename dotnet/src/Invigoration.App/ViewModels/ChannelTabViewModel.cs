using System.Collections.ObjectModel;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.App.Models;
using Invigoration.Core.Chat;
using Stimpak;

namespace Invigoration.App.ViewModels;

/// <summary>
/// One joined SC2/SC:R/WC3:R channel, shown as its own sub-tab within a bot's
/// tab — see BotTabViewModel.Channels. Its own chat log, and its own roster
/// bound straight to Stimpak's own PeopleRegistry.Channel(index) (handed
/// back by BotEngine.Sc2ChannelJoined) rather than a reconciled-in-App-code
/// copy — that roster is already correct per-channel, so there's nothing to
/// duplicate.
/// </summary>
public sealed partial class ChannelTabViewModel(byte channelIndex, ChatChannel channel, ObservableCollection<Person> users) : ViewModelBase
{
    public byte ChannelIndex { get; } = channelIndex;

    public string Title { get; } = channel.Name;

    public ObservableCollection<ChatLineViewModel> ChatLines { get; } = [];

    public ObservableCollection<Person> Users { get; } = users;

    /// <summary>Set by BotTabViewModel.OnChatMessage when a new message arrives here while this isn't the bot's SelectedChannel — cleared when it becomes selected (OnSelectedChannelChanged).</summary>
    [ObservableProperty]
    public partial bool HasUnread { get; set; }

    /// <summary>
    /// Trimmed version of BotTabViewModel.HandleChatEvent: only the branches
    /// that render a chat-log line. No UpsertUser/roster mutation at all —
    /// Users already updates itself reactively (Stimpak's own PeopleRegistry
    /// is fed unconditionally in BotEngine.Sc2.cs regardless of which
    /// channel an event targets), so there's no App-side roster-tracking
    /// machinery to write here, unlike the classic-BNCS ChannelUserViewModel
    /// path this deliberately does not reuse. No Whisper/WhisperSent case
    /// either — a whisper's ChannelIndex is always null (not channel-scoped),
    /// so BotTabViewModel.OnChatMessage never routes one here at all; see
    /// BotTabViewModel.WhisperThreads for where whispers actually go.
    /// </summary>
    public void HandleChatEvent(ChatEvent e, ChatPalette palette, Bitmap? userIcon = null)
    {
        switch (e.Type)
        {
            case ChatEventType.Join:
                ChatLines.Add(new ChatLineViewModel($"*** {e.Username} has joined the channel.", palette.Gray));
                break;

            case ChatEventType.Leave:
                ChatLines.Add(new ChatLineViewModel($"*** {e.Username} has left the channel.", palette.Gray));
                break;

            case ChatEventType.Talk:
                var segments = new List<ChatLogSegment> { new(palette.GetUserNameColor(e.Flags), $"{e.Username}: ") };
                segments.AddRange(ChatColorFormatter.Parse(e.Text, palette.GetChatColor(e.Flags), palette));
                ChatLines.Add(new ChatLineViewModel(segments, userIcon));
                break;

            case ChatEventType.Emote:
                ChatLines.Add(new ChatLineViewModel($"<{e.Username} {e.Text}>", palette.GetEmoteColor(e.Flags), userIcon));
                break;

            case ChatEventType.Info:
                ChatLines.Add(new ChatLineViewModel(e.Text, palette.Info));
                break;

            case ChatEventType.Error:
                ChatLines.Add(new ChatLineViewModel(e.Text, palette.Error));
                break;

            case ChatEventType.Broadcast:
                ChatLines.Add(new ChatLineViewModel($"[Broadcast]: {e.Text}", palette.Debug));
                break;
        }
    }
}
