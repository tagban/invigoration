using System.Collections.ObjectModel;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.App.Models;
using Invigoration.Core.Chat;
using Invigoration.Core.Sc2;
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

    /// <summary>Fired after ChatLineTrimmer trims old lines off ChatLines — see BotTabViewModel.ChatLinesTrimmed's matching remarks (this channel's own analog, for a busy long-lived SC2/SC:R/WC3:R channel's chat log). ChannelTabView rebuilds its Inlines from the (now bounded) collection in response.</summary>
    public event Action? ChatLinesTrimmed;

    /// <summary>
    /// Must be called once, right after construction (see BotTabViewModel.OnSc2ChannelJoined,
    /// the only place this class is instantiated) — a primary constructor has no ordinary body,
    /// and a field initializer can't reference any instance member at all (CS0236, not just
    /// method calls, as an earlier attempt here assumed), so there's no way to wire this up
    /// during construction itself.
    /// </summary>
    public void AttachChatLineTrimmer() => ChatLineTrimmer.Attach(ChatLines, () => ChatLinesTrimmed?.Invoke());

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
    /// <summary>The dim gray a name's "#1234" code is shown in.</summary>
    private static readonly RgbColor NameCodeColor = new(0x6A, 0x6A, 0x6A);

    /// <param name="hideNameCodes">Leave out a name's "#123" code, as StarCraft II's own chat does.</param>
    public void HandleChatEvent(ChatEvent e, ChatPalette palette, Bitmap? userIcon = null, bool showIcons = true, bool largePictures = false, bool hideNameCodes = false)
    {
        switch (e.Type)
        {
            case ChatEventType.Join:
                ChatLines.Add(new ChatLineViewModel($"*** {e.Username} has joined the channel.", palette.Gray));
                break;

            case ChatEventType.Leave:
                ChatLines.Add(new ChatLineViewModel($"*** {e.Username} has left the channel.", palette.Gray));
                break;

            case ChatEventType.Talk when Invigoration.Core.Discord.DiscordRelayLine.TryParse(e.Text, out var relayedUser, out var relayedText):
                // Another Invigoration bot's Discord relay — shown under the Discord user's name with the logo, same as a classic bot tab.
                ChatLines.Add(DiscordRelayRendering.Build(relayedUser, relayedText, relayedBy: e.Username, palette, showIcons));
                break;

            case ChatEventType.Talk:
                // "Name#1234": the code dimmed, so it's there but reads like a chat.
                var nameColor = palette.GetUserNameColor(e.Flags);
                var (name, code) = NameParts.Split(e.Username);
                var segments = new List<ChatLogSegment> { new(nameColor, name) };
                if (code.Length > 0 && !hideNameCodes)
                {
                    segments.Add(new(NameCodeColor, code));
                }

                segments.Add(new(nameColor, ": "));
                var nameSegments = segments.Count;
                segments.AddRange(ChatColorFormatter.Parse(e.Text, palette.GetChatColor(e.Flags), palette));
                ChatLines.Add(new ChatLineViewModel(segments, userIcon, largePictures, nameSegments)
                {
                    ClanTag = largePictures ? NativeMemberPortraits.ClanTagFor(e.Username) : "",
                });
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
