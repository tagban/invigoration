using System.Runtime.InteropServices;
using Invigoration.Core.Chat;
using Invigoration.Core.Clan;
using Invigoration.Core.Crypto;
using Invigoration.Core.Protocol;
using Invigoration.Core.Text;

namespace Invigoration.Core;

/// <summary>Port of modCommands.bas's ParseCommand.</summary>
public sealed partial class BotEngine
{
    /// <summary>
    /// No URL here on purpose — Battle.net/PVPGN chat filters can mangle
    /// "https://..." text (e.g. rendering it as "https://github!@#$/"), so
    /// the version reply just states the build instead of linking out.
    /// </summary>
    private const string VersionLine = $"/me is an Invigoration v{AppVersion.Current}";

    /// <summary>Runs a command typed locally in the bot's own UI — always trusted, no bot-master check.</summary>
    public Task RunLocalCommandAsync(string message) => HandleCommandAsync(Config.Username, message, isLocal: true, isWhisper: false);

    private Task HandleCommandAsync(string username, string message, bool isWhisper) =>
        HandleCommandAsync(username, message, isLocal: false, isWhisper);

    private async Task HandleCommandAsync(string username, string message, bool isLocal, bool isWhisper)
    {
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

        if (!isLocal && !IsAuthorized(username, command, rest))
        {
            return;
        }

        Task Reply(string text) => ReplyAsync(text, username, isWhisper);

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
                await Reply(VersionLine).ConfigureAwait(false);
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

            case "clanadd":
                await HandleClanAddAsync(rest, Reply).ConfigureAwait(false);
                break;

            case "clanremove":
                await HandleClanRemoveAsync(rest, Reply).ConfigureAwait(false);
                break;

            case "clanrank":
                await HandleClanRankAsync(rest, Reply).ConfigureAwait(false);
                break;

            case "clanalias":
                await HandleClanAliasAsync(rest, Reply, add: true).ConfigureAwait(false);
                break;

            case "clanunalias":
                await HandleClanAliasAsync(rest, Reply, add: false).ConfigureAwait(false);
                break;

            case "claninfo":
                await HandleClanInfoAsync(rest, Reply).ConfigureAwait(false);
                break;

            case "clanlist":
                await HandleClanListAsync(rest, Reply).ConfigureAwait(false);
                break;

            case "clanscore":
                await HandleClanScoreAsync(rest, Reply).ConfigureAwait(false);
                break;

            case "trivia":
                await HandleTriviaCommandAsync(rest, username, Reply).ConfigureAwait(false);
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

    /// <summary>
    /// The bot master always has full access. A member whose clan rank
    /// matches <see cref="Config"/>'s BannedRank is blocked from everything
    /// else, including "trivia join"/"trivia score"/"trivia categories"
    /// below, regardless of any other grant. Otherwise "trivia join"/
    /// "trivia score"/"trivia categories" are always open — gating trivia's
    /// own entry point (and the harmless, read-only category listing) behind
    /// the same allowlist as admin commands like "kick"/"setpass" would make
    /// it unplayable for anyone the bot master hasn't individually granted
    /// access to ("trivia on"/"trivia off"/"trivia &lt;category&gt;", round
    /// control, still go through the normal check below).
    /// Beyond that, a user is authorized if the clan rank they resolve to via
    /// the shared roster (matched by their primary name or a tracked alias)
    /// has the resolved canonical command in its <see cref="Clan.ClanRank.AllowedCommands"/>
    /// — aliases like "h"/"hex" share one grant via
    /// <see cref="Commands.CommandCatalog.ResolveCanonicalName"/>. There's no
    /// longer a separate per-user grant list (the old PermissionLevel system)
    /// — access lives entirely on the rank now, shared across every bot.
    ///
    /// Every identity comparison here goes through
    /// <see cref="BnetUsername.MatchesOnServer"/> (not the plain
    /// <see cref="BnetUsername.Equals"/>), scoped to this bot's own
    /// Config.BattlenetServer — a Diablo II player showing as "*Name"
    /// in-game still matches (same as before), but a same-named account on
    /// a DIFFERENT Battle.net server no longer does, unless the BotMaster/
    /// alias entry was deliberately left unqualified. Classic accounts are
    /// scoped per gateway: "tagban" on useast.battle.net and "tagban" on
    /// asia.battle.net are unrelated accounts that happen to share a name,
    /// not the same person — without this, someone registering the right
    /// name on a different server the shared roster also happens to
    /// reference could be handed bot-master access or a trusted rank they
    /// were never actually granted.
    /// </summary>
    private bool IsAuthorized(string username, string typedCommand, string rest = "")
    {
        if (BnetUsername.MatchesOnServer(username, Config.BotMaster, Config.BattlenetServer))
        {
            return true;
        }

        if (IsBannedUser(username))
        {
            return false;
        }

        if (typedCommand.Equals("trivia", StringComparison.OrdinalIgnoreCase) &&
            (rest.Equals("score", StringComparison.OrdinalIgnoreCase) ||
             rest.Equals("join", StringComparison.OrdinalIgnoreCase) ||
             rest.Equals("categories", StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        var canonical = Commands.CommandCatalog.ResolveCanonicalName(typedCommand);
        var rankName = ClanRosterStore.FindTrusted(username, Config.BattlenetServer)?.Rank;
        if (string.IsNullOrEmpty(rankName))
        {
            return false;
        }

        var rank = ClanRankStore.Find(rankName);
        return rank is not null && rank.AllowedCommands.Any(c => c.Equals(canonical, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>True if this username resolves to a tracked member whose rank matches Config.BannedRank — used both to block commands (IsAuthorized) and to stop trivia from scoring/reacting to their chat at all. Server-scoped (see FindTrusted) so a same-named account on another server can't inherit someone else's ban, nor evade their own by being looked up under the wrong server's identity.</summary>
    private bool IsBannedUser(string username)
    {
        var rank = ClanRosterStore.FindTrusted(username, Config.BattlenetServer)?.Rank;
        return !string.IsNullOrEmpty(rank) && !string.IsNullOrEmpty(Config.BannedRank) &&
            rank.Equals(Config.BannedRank, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// "clanadd &lt;name&gt; [rank]" — adds a member, or updates their rank if
    /// already tracked (matched against this bot's own server via
    /// FindTrusted, so it never touches an unrelated same-named account on a
    /// different server). A freshly-added member gets the typed name as
    /// their nickname and this bot's own server auto-appended to their
    /// primary account — "fusion" typed on a useast bot becomes
    /// Name="fusion@useast.battle.net", NickName="fusion" — safe by default
    /// without the caller needing to type the "name@server" syntax themselves.
    /// </summary>
    private Task HandleClanAddAsync(string rest, Func<string, Task> reply)
    {
        var parts = rest.Split(' ', 2);
        var name = parts[0];
        if (name.Length == 0)
        {
            return reply("Usage: clanadd <name> [rank]");
        }

        var rank = parts.Length > 1 ? parts[1].Trim() : "";
        var member = ClanRosterStore.FindTrusted(name, Config.BattlenetServer);
        if (member is null)
        {
            var qualifiedName = $"{name}@{Config.BattlenetServer}";
            member = new ClanMember { Name = qualifiedName, NickName = name, Rank = rank, IsClanMember = true };
            ClanRosterStore.Members.Add(member);
            ClanRosterStore.Save();
            return reply(rank.Length > 0
                ? $"Added {name} ({qualifiedName}) to the clan roster as \"{rank}\"."
                : $"Added {name} ({qualifiedName}) to the clan roster.");
        }

        member.Rank = rank;
        member.IsClanMember = true;
        ClanRosterStore.Save();
        return reply(rank.Length > 0 ? $"Added {name} to the clan roster as \"{rank}\"." : $"Added {name} to the clan roster.");
    }

    private Task HandleClanRemoveAsync(string rest, Func<string, Task> reply)
    {
        var member = ClanRosterStore.FindTrusted(rest, Config.BattlenetServer);
        if (member is null)
        {
            return reply($"{rest} isn't in the clan roster.");
        }

        ClanRosterStore.Members.Remove(member);
        ClanRosterStore.Save();
        return reply($"Removed {member.Name} from the clan roster.");
    }

    /// <summary>"clanrank &lt;name&gt; &lt;rank&gt;"</summary>
    private Task HandleClanRankAsync(string rest, Func<string, Task> reply)
    {
        var parts = rest.Split(' ', 2);
        if (parts.Length < 2)
        {
            return reply("Usage: clanrank <name> <rank>");
        }

        var member = ClanRosterStore.FindTrusted(parts[0], Config.BattlenetServer);
        if (member is null)
        {
            return reply($"{parts[0]} isn't in the clan roster.");
        }

        member.Rank = parts[1].Trim();
        ClanRosterStore.Save();
        return reply($"{member.Name} is now ranked \"{member.Rank}\".");
    }

    /// <summary>"clanalias &lt;name&gt; &lt;alias&gt;" / "clanunalias &lt;name&gt; &lt;alias&gt;"</summary>
    private Task HandleClanAliasAsync(string rest, Func<string, Task> reply, bool add)
    {
        var parts = rest.Split(' ', 2);
        if (parts.Length < 2)
        {
            return reply(add ? "Usage: clanalias <name> <alias>" : "Usage: clanunalias <name> <alias>");
        }

        var member = ClanRosterStore.FindTrusted(parts[0], Config.BattlenetServer);
        if (member is null)
        {
            return reply($"{parts[0]} isn't in the clan roster.");
        }

        var alias = parts[1].Trim();
        if (add)
        {
            if (!member.Aliases.Any(a => a.Equals(alias, StringComparison.OrdinalIgnoreCase)))
            {
                member.Aliases.Add(alias);
            }
        }
        else
        {
            member.Aliases.RemoveAll(a => a.Equals(alias, StringComparison.OrdinalIgnoreCase));
        }

        ClanRosterStore.Save();
        return reply(add ? $"{member.Name} may also be seen as {alias}." : $"Removed {alias} from {member.Name}'s aliases.");
    }

    private Task HandleClanInfoAsync(string rest, Func<string, Task> reply)
    {
        var member = ClanRosterStore.FindTrusted(rest, Config.BattlenetServer);
        if (member is null)
        {
            return reply($"{rest} isn't in the clan roster.");
        }

        var aliases = member.Aliases.Count > 0 ? string.Join(", ", member.Aliases) : "none";
        var rank = member.Rank.Length > 0 ? member.Rank : "unranked";
        var lastSeen = member.LastSeenUtc is { } seenUtc ? FormatLastSeen(seenUtc) : "never";
        return reply(
            $"{member.Name} :: rank: {rank} :: aliases: {aliases} :: trivia score: {member.TriviaScore} :: last seen: {lastSeen}");
    }

    /// <summary>"clanscore &lt;name&gt; &lt;+/-delta&gt;" — adjusts a tracked member's trivia score; the running total a future trivia game would add to.</summary>
    private Task HandleClanScoreAsync(string rest, Func<string, Task> reply)
    {
        var parts = rest.Split(' ', 2);
        if (parts.Length < 2 ||
            !int.TryParse(parts[1], System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out var delta))
        {
            return reply("Usage: clanscore <name> <+/-delta>");
        }

        var member = ClanRosterStore.FindTrusted(parts[0], Config.BattlenetServer);
        if (member is null)
        {
            return reply($"{parts[0]} isn't in the clan roster.");
        }

        member.TriviaScore += delta;
        ClanRosterStore.Save();
        return reply($"{member.Name}'s trivia score is now {member.TriviaScore}.");
    }

    private static string FormatLastSeen(DateTime seenUtc)
    {
        var span = DateTime.UtcNow - seenUtc;
        if (span < TimeSpan.FromMinutes(1))
        {
            return "just now";
        }

        if (span < TimeSpan.FromHours(1))
        {
            return $"{(int)span.TotalMinutes}m ago";
        }

        if (span < TimeSpan.FromDays(1))
        {
            return $"{(int)span.TotalHours}h ago";
        }

        return $"{(int)span.TotalDays}d ago";
    }

    /// <summary>"clanlist [rank]" — lists everyone, or just members of the given rank.</summary>
    private Task HandleClanListAsync(string rest, Func<string, Task> reply)
    {
        var members = rest.Length == 0
            ? ClanRosterStore.Members
            : ClanRosterStore.Members.Where(m => m.Rank.Equals(rest, StringComparison.OrdinalIgnoreCase)).ToList();

        if (members.Count == 0)
        {
            return reply(rest.Length == 0 ? "The clan roster is empty." : $"No members ranked \"{rest}\".");
        }

        return reply(string.Join(", ", members.Select(m => m.Rank.Length > 0 ? $"{m.Name} ({m.Rank})" : m.Name)));
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
            "ver" => VersionLine,
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
