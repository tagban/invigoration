using System.Collections.ObjectModel;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.Models;
using Invigoration.Core;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;

namespace Invigoration.App.ViewModels;

/// <summary>One bot tab: wraps a BotEngine and projects its events onto observable collections for binding.</summary>
public partial class BotTabViewModel : ViewModelBase, IAsyncDisposable
{
    public BotEngine Engine { get; }

    public BotConfig Config => Engine.Config;

    public string Title => Config.DisplayName;

    public ObservableCollection<ChatLineViewModel> ChatLines { get; } = [];

    public ObservableCollection<ChannelUserViewModel> ChannelUsers { get; } = [];

    [ObservableProperty]
    public partial string InputText { get; set; } = "";

    [ObservableProperty]
    public partial bool IsConnected { get; set; }

    [ObservableProperty]
    public partial string StatusText { get; set; } = "Disconnected";

    public bool DebugMode
    {
        get => Engine.DebugMode;
        set => Engine.DebugMode = value;
    }

    public BotTabViewModel(BotEngine engine)
    {
        Engine = engine;
        Engine.Log += OnLog;
        Engine.ChatMessage += OnChatMessage;
        Engine.BncsConnected += () => Dispatcher.UIThread.Post(() =>
        {
            IsConnected = true;
            StatusText = "Connected";
        });
        Engine.BncsDisconnected += _ => Dispatcher.UIThread.Post(() =>
        {
            IsConnected = false;
            StatusText = "Disconnected";
            ChannelUsers.Clear();
        });
    }

    [RelayCommand]
    private async Task ConnectAsync()
    {
        try
        {
            StatusText = "Connecting...";
            await Engine.ConnectAsync();
        }
        catch (Exception ex)
        {
            ChatLines.Add(new ChatLineViewModel($"Connect failed: {ex.Message}", ChatColors.Red));
            StatusText = "Disconnected";
        }
    }

    [RelayCommand]
    private Task DisconnectAsync() => Engine.DisconnectAsync();

    [RelayCommand]
    private async Task SendAsync()
    {
        var text = InputText;
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        InputText = "";

        if (text.Length > 0 && (text[0] == Config.Trigger.FirstOrDefault() || text[0] == '/'))
        {
            await Engine.RunLocalCommandAsync(text);
        }
        else
        {
            await Engine.SendChatCommandAsync(text);
            ChatLines.Add(new ChatLineViewModel(BuildUserLine("me", text)));
        }
    }

    private void OnLog(IReadOnlyList<ChatLogSegment> segments) =>
        Dispatcher.UIThread.Post(() => ChatLines.Add(new ChatLineViewModel(segments)));

    private void OnChatMessage(ChatEvent e) => Dispatcher.UIThread.Post(() => HandleChatEvent(e));

    private void HandleChatEvent(ChatEvent e)
    {
        switch (e.Type)
        {
            case ChatEventType.Channel:
                ChannelUsers.Clear();
                ChatLines.Add(new ChatLineViewModel($"*** Joined channel: {e.Text}", ChatColors.LtBlue));
                break;

            case ChatEventType.ShowUser:
            case ChatEventType.Join:
                UpsertUser(e);
                if (e.Type == ChatEventType.Join)
                {
                    ChatLines.Add(new ChatLineViewModel($"*** {e.Username} has joined the channel.", ChatColors.Gray));
                }

                break;

            case ChatEventType.Leave:
                var leaving = ChannelUsers.FirstOrDefault(u => u.Username == e.Username);
                if (leaving is not null)
                {
                    ChannelUsers.Remove(leaving);
                }

                ChatLines.Add(new ChatLineViewModel($"*** {e.Username} has left the channel.", ChatColors.Gray));
                break;

            case ChatEventType.UserFlags:
                UpsertUser(e);
                break;

            case ChatEventType.Talk:
                ChatLines.Add(new ChatLineViewModel(BuildUserLine(e.Username, e.Text)));
                break;

            case ChatEventType.Emote:
                ChatLines.Add(new ChatLineViewModel($"<{e.Username} {e.Text}>", ChatColors.Purple));
                break;

            case ChatEventType.Whisper:
                ChatLines.Add(new ChatLineViewModel($"[{e.Username} whispers]: {e.Text}", ChatColors.MedBlue));
                break;

            case ChatEventType.WhisperSent:
                ChatLines.Add(new ChatLineViewModel($"[whisper to {e.Username}]: {e.Text}", ChatColors.MedBlue));
                break;

            case ChatEventType.Info:
                ChatLines.Add(new ChatLineViewModel(e.Text, ChatColors.White));
                break;

            case ChatEventType.Error:
                ChatLines.Add(new ChatLineViewModel(e.Text, ChatColors.Red));
                break;

            case ChatEventType.Broadcast:
                ChatLines.Add(new ChatLineViewModel($"[Broadcast]: {e.Text}", ChatColors.Orange));
                break;
        }
    }

    private void UpsertUser(ChatEvent e)
    {
        var user = ChannelUsers.FirstOrDefault(u => u.Username == e.Username);
        if (user is null)
        {
            user = new ChannelUserViewModel(e.Username);
            ChannelUsers.Add(user);
        }

        user.Flags = e.Flags;
        user.Ping = e.Ping;
        if (e.Text.Length > 0)
        {
            user.StatString = e.Text;
        }
    }

    private static IReadOnlyList<ChatLogSegment> BuildUserLine(string username, string text)
    {
        var segments = new List<ChatLogSegment> { new(ChatColors.LtBlue, $"{username}: ") };
        segments.AddRange(ChatColorFormatter.Parse(text, ChatColors.White));
        return segments;
    }

    public ValueTask DisposeAsync()
    {
        Engine.Log -= OnLog;
        Engine.ChatMessage -= OnChatMessage;
        return Engine.DisposeAsync();
    }
}
