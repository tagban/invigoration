using System.Runtime.InteropServices;
using Invigoration.Core.Chat;
using Invigoration.Core.Crypto;
using Invigoration.Core.Protocol;
using Invigoration.Core.Text;

namespace Invigoration.Core;

/// <summary>Port of modCommands.bas's ParseCommand.</summary>
public sealed partial class BotEngine
{
    private const string BotVersion = "2.0.0-dotnet";

    /// <summary>Runs a command typed locally in the bot's own UI — always trusted, no bot-master check.</summary>
    public Task RunLocalCommandAsync(string message) => HandleCommandAsync(Config.Username, message, isLocal: true);

    private Task HandleCommandAsync(string username, string message) =>
        HandleCommandAsync(username, message, isLocal: false);

    private async Task HandleCommandAsync(string username, string message, bool isLocal)
    {
        var isWhisper = !isLocal; // network-triggered commands only arrive via whisper/talk; talk-triggered ones reply publicly too, see below
        if (message.Equals("?trigger", StringComparison.OrdinalIgnoreCase))
        {
            message = Config.Trigger + "trigger";
        }

        if (message.Length == 0 || (message[0] != Config.Trigger[0] && message[0] != '/'))
        {
            return;
        }

        message = message[1..];

        var parts = message.Split(' ', 2);
        var command = parts[0];
        var rest = parts.Length > 1 ? parts[1].Trim() : "";

        if (!isLocal && !username.Equals(Config.BotMaster, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        Task Reply(string text) => ReplyAsync(text, username, isWhisper && !isLocal);

        switch (command.ToLowerInvariant())
        {
            case "idle":
                await HandleIdleCommandAsync(rest, Reply).ConfigureAwait(false);
                break;

            case "disconnect":
            case "disc":
                LogInfo("Disconnecting...");
                await DisconnectAsync().ConfigureAwait(false);
                break;

            case "colors":
            case "color":
                LogInfo("Chat Colors Help:");
                LogInfo("Use a non-breaking space (Alt+0160) then a letter: r=red w=white q=gray g=green y=yellow b=blue o=orange c=cyan p=purple l=light-yellow e=beige k=pink");
                break;

            case "reconnect":
                LogInfo("Reconnecting, hold on tight!");
                await DisconnectAsync().ConfigureAwait(false);
                await ConnectAsync().ConfigureAwait(false);
                break;

            case "hex":
            case "h":
                await SendChatCommandAsync("£" + HexCodec.StrToHex(rest)).ConfigureAwait(false);
                break;

            case "invigencrypt":
            case "encrypt":
            case "ie":
            case "i":
                await SendChatCommandAsync("" + InvigCipher.Encrypt(rest + "-")).ConfigureAwait(false);
                break;

            case "sysinfo":
                await Reply($"Invigoration running on: {RuntimeInformation.OSDescription}.").ConfigureAwait(false);
                break;

            case "ver":
                await Reply($"/me is an Invigoration v{BotVersion} - https://github.com/").ConfigureAwait(false);
                break;

            case "uptime":
                await Reply($"/me has been online for: {FormatUptime()}.").ConfigureAwait(false);
                break;

            case "about":
                await Reply("Invigoration, originally written in Visual Basic by Tagban since 2004; ported to C#/.NET.")
                    .ConfigureAwait(false);
                break;

            case "say":
                await SendChatCommandAsync(rest).ConfigureAwait(false);
                break;

            case "bancount":
                await Reply(FormatCount(_session.BanCount, "banned")).ConfigureAwait(false);
                break;

            case "kickcount":
                await Reply(FormatCount(_session.KickCount, "kicked")).ConfigureAwait(false);
                break;

            case "joincount":
                await Reply(FormatCount(_session.JoinCount, "joined the channel")).ConfigureAwait(false);
                break;

            case "ban":
                await SendChatCommandAsync($"/ban {rest}").ConfigureAwait(false);
                break;

            case "kick":
                await SendChatCommandAsync($"/kick {rest}").ConfigureAwait(false);
                break;

            case "join":
                await SendBncsAsync(new PacketWriter(), BncsPacketId.SID_LEAVECHAT).ConfigureAwait(false);
                await SendBncsAsync(new PacketWriter().WriteDword(2).WriteNTString(rest), BncsPacketId.SID_JOINCHANNEL)
                    .ConfigureAwait(false);
                break;

            case "user":
                _session.TargetUser = rest;
                break;

            case "useroff":
                _session.TargetUser = "";
                break;

            case "prepend":
            case "pre":
                _session.PrependText = rest;
                LogInfo($"\"{rest}\" will be shown before each send.");
                break;

            case "postpend":
            case "post":
                _session.PostpendText = rest;
                LogInfo($"\"{rest}\" will be shown after each send.");
                break;

            case "setmaster":
                Config.BotMaster = rest;
                await Reply("Bot master changed!").ConfigureAwait(false);
                break;

            case "sethome":
                Config.HomeChannel = rest;
                await Reply("Home channel changed!").ConfigureAwait(false);
                break;

            case "setusername":
                Config.Username = rest;
                await Reply("Login username changed!").ConfigureAwait(false);
                break;

            case "setpass":
                Config.Password = rest;
                await Reply("Login password changed!").ConfigureAwait(false);
                break;

            case "setserver":
                Config.BattlenetServer = rest;
                await Reply("Server changed!").ConfigureAwait(false);
                break;

            case "settrigger":
                Config.Trigger = rest;
                await Reply("Bot trigger changed!").ConfigureAwait(false);
                break;

            case "trigger":
                await Reply($"The bot's trigger is: {Config.Trigger}").ConfigureAwait(false);
                break;

            case "lastm":
            case "lastw":
            case "last":
            case "lrm":
            case "lrw":
                await Reply($"Last whisper received from: {_session.LastWhisperFromUser} :: {_session.LastWhisperFromText}")
                    .ConfigureAwait(false);
                break;

            case "lastsm":
            case "lastsw":
            case "lastsend":
            case "lsm":
            case "lsw":
                await Reply($"Last whisper sent to: {_session.LastWhisperSentUser} :: {_session.LastWhisperSentText}")
                    .ConfigureAwait(false);
                break;

            case "canada":
                _session.CanadaMode = !_session.CanadaMode;
                await Reply($"Canada Mode {(_session.CanadaMode ? "enabled" : "disabled")}.").ConfigureAwait(false);
                break;

            case "accept":
                _session.AcceptClanInvites = !_session.AcceptClanInvites;
                await Reply($"Clan invite auto-accept {(_session.AcceptClanInvites ? "enabled" : "disabled")}.")
                    .ConfigureAwait(false);
                break;

            case "debug":
                _session.DebugMode = !_session.DebugMode;
                LogInfo($"Debug mode {(_session.DebugMode ? "enabled" : "disabled")}.");
                break;

            case "leetspeak":
                _session.LeetSpeakMode = !_session.LeetSpeakMode;
                await Reply($"Leet speak {(_session.LeetSpeakMode ? "enabled" : "disabled")}.").ConfigureAwait(false);
                break;

            case "fudd":
                _session.FuddMode = !_session.FuddMode;
                await Reply($"Elmer Fudd mode {(_session.FuddMode ? "enabled" : "disabled")}.").ConfigureAwait(false);
                break;

            case "moo":
                _session.MooMode = !_session.MooMode;
                await Reply(_session.MooMode ? "Moooooooooooooooo mode engaged!" : "Cows are off...")
                    .ConfigureAwait(false);
                break;

            case "home":
            case "gohome":
            case "homechan":
            case "homechannel":
                await JoinHomeAsync().ConfigureAwait(false);
                break;

            default:
                // Everything else (whois, /w, /friends, etc.) is a raw server
                // command only relayed when typed locally by the operator,
                // matching the original's Inbot-only passthrough.
                if (isLocal)
                {
                    await SendChatCommandAsync("/" + message).ConfigureAwait(false);
                }

                break;
        }
    }

    private Task HandleIdleCommandAsync(string rest, Func<string, Task> reply)
    {
        if (rest.Equals("off", StringComparison.OrdinalIgnoreCase))
        {
            _session.IdleTimeSetMinutes = 0;
            _session.IdleMessage = "";
            return reply("Idle message turned off.");
        }

        var parts = rest.Split(' ', 2);
        if (!int.TryParse(parts[0], out var minutes) || parts.Length < 2)
        {
            return Task.CompletedTask;
        }

        _session.IdleTimeSetMinutes = minutes;
        _session.IdleMessage = parts[1].ToLowerInvariant() switch
        {
            "uptime" => $"/me has been online for: {FormatUptime()}.",
            "ver" => $"/me is an Invigoration v{BotVersion} - https://github.com/",
            _ => parts[1],
        };

        return reply("Idle message set.");
    }

    private Task ReplyAsync(string text, string username, bool asWhisper) =>
        asWhisper ? SendChatCommandAsync($"/w {username} {text}") : SendChatCommandAsync(text);

    private static string FormatCount(int count, string verb) => count switch
    {
        0 => $"No one has {verb} since I joined this channel.",
        1 => $"1 user has {verb} since I joined this channel.",
        _ => $"{count} users have {verb} since I joined this channel.",
    };

    private string FormatUptime() => (DateTimeOffset.UtcNow - _connectedAt).ToString(@"d\d\ hh\h\ mm\m\ ss\s");
}
