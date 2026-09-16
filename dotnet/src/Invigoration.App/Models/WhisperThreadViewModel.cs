using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.ViewModels;

namespace Invigoration.App.Models;

/// <summary>
/// One whisper conversation with a single peer, from one bot's point of view — built up from
/// ChatEventType.Whisper (incoming) and WhisperSent (outgoing) events by
/// BotTabViewModel.UpsertWhisper. Owner is set once at creation and never changes; it's what
/// lets the global cross-bot Whispers tab (MainWindowViewModel) label each thread by login and
/// route a reply through the right bot's engine, while the exact same instance is also what
/// that bot's own per-bot Whispers tab shows — replying from either place updates one thread.
/// </summary>
/// <summary>
/// Whatever a whisper thread belongs to — a Battle.net bot, or a Hotline session. The global
/// Whispers tab shows threads from both, so it needs a name to label them by and a way to send a
/// reply back through whichever connection the thread came in on.
/// </summary>
public interface IWhisperHost
{
    /// <summary>What to call this connection in the Whispers tab ("via ...").</summary>
    string Title { get; }

    /// <summary>Sends this thread's DraftText to its peer and clears it.</summary>
    Task SendWhisperAsync(WhisperThreadViewModel thread);
}

public sealed partial class WhisperThreadViewModel(IWhisperHost owner, string peer) : ObservableObject
{
    public IWhisperHost Owner { get; } = owner;

    /// <summary>
    /// Sending lives on the thread rather than on the owner so the view binds to one command
    /// regardless of which kind of connection the thread belongs to — a bot's whisper and a
    /// Hotline private message reach the same box.
    /// </summary>
    [RelayCommand]
    private Task Send() => Owner.SendWhisperAsync(this);

    public string Peer { get; } = peer;

    public ObservableCollection<ChatLineViewModel> Messages { get; } = [];

    [ObservableProperty]
    public partial DateTime LastActivityUtc { get; set; }

    /// <summary>True from the moment an incoming whisper arrives until this thread is selected in some Whispers tab (per-bot or global) — see BotTabViewModel.UpsertWhisper / MarkRead.</summary>
    [ObservableProperty]
    public partial bool HasUnread { get; set; }

    /// <summary>Bound to the reply textbox — both the per-bot Whispers tab and the friends-list right-click compose popup write into this.</summary>
    [ObservableProperty]
    public partial string DraftText { get; set; } = "";

    public void MarkRead() => HasUnread = false;
}
