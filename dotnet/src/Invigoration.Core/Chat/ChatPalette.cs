namespace Invigoration.Core.Chat;

/// <summary>
/// A named chat-log color scheme, ported from github.com/tagban/bnubot's
/// net.bnubot.bot.gui.colors package (ColorScheme + StarcraftColorScheme +
/// Diablo2ColorScheme + InvigorationColorScheme) — this bot's own author's
/// prior Java client, given as the reference to use here rather than
/// invented values. Field names and the three Get*Color methods mirror
/// bnubot's ColorScheme role names and flag-priority branches exactly
/// (getUserNameColor/getChatColor/getEmoteColor); <see cref="UserFlags"/>'s
/// bit values already line up with bnubot's raw flag constants (Blizzard
/// =BLIZZARD_REP, Operator, Speaker, Admin=BNET_REP, Squelched=ignored,
/// Special=BLIZZARD_GUEST). <see cref="Background"/> and
/// <see cref="Highlight"/> have no bnubot equivalent — Background because
/// bnubot never asked its author to make one legible against pure-blue text
/// (StarCraft's info color on a pure-black background is exactly that
/// problem), Highlight because it's this port's own addition for marking
/// the active bot tab / a focused chat input.
/// </summary>
public sealed class ChatPalette
{
    // --- Single-value roles ---

    public required RgbColor Background { get; init; }

    /// <summary>Default chat/message body text (bnubot's ChatColor/EmoteColor "PRIORITY_NORMAL" default, and getForegroundColor for StarCraft/Invig).</summary>
    public required RgbColor White { get; init; }

    /// <summary>"*** Joined channel: X" (bnubot's ChannelColor).</summary>
    public required RgbColor Channel { get; init; }

    /// <summary>System/status messages (bnubot's InfoColor).</summary>
    public required RgbColor Info { get; init; }

    /// <summary>Errors and warnings (bnubot's ErrorColor).</summary>
    public required RgbColor Error { get; init; }

    /// <summary>Debug/hex-dump log lines (bnubot's DebugColor).</summary>
    public required RgbColor Debug { get; init; }

    /// <summary>Join/leave notices — bnubot has no dedicated role for these; each scheme picks its own dim tone.</summary>
    public required RgbColor Gray { get; init; }

    /// <summary>The bot's own name prefix on its outgoing chat lines (bnubot's SelfUserNameColor).</summary>
    public required RgbColor SelfUserName { get; init; }

    /// <summary>Whisper lines, and also the "ignored user" chat-text tone in all three source schemes (bnubot's WhisperColor).</summary>
    public required RgbColor Whisper { get; init; }

    /// <summary>Marks the active bot tab and/or a focused chat-input box. Not part of bnubot's palette.</summary>
    public required RgbColor Highlight { get; init; }

    // --- Flag-priority branch colors (shared across UserName/Chat/Emote; exposed mainly for the Get*Color methods below) ---

    public required RgbColor Red { get; init; }
    public required RgbColor Green { get; init; }
    public required RgbColor Cyan { get; init; }
    public required RgbColor Speaker { get; init; }
    public required RgbColor Guest { get; init; }
    public required RgbColor UserNameDefault { get; init; }
    public required RgbColor EmoteDefault { get; init; }

    /// <summary>Ported from bnubot's getUserNameColor(flags): colors a user's name by their highest-priority channel flag.</summary>
    public RgbColor GetUserNameColor(uint flags)
    {
        var f = (UserFlags)flags;
        if (f.HasFlag(UserFlags.Squelched)) return Red;
        if (f.HasFlag(UserFlags.Blizzard)) return Cyan;
        if (f.HasFlag(UserFlags.Admin)) return Green;
        if (f.HasFlag(UserFlags.Operator)) return White;
        if (f.HasFlag(UserFlags.Speaker)) return Speaker;
        if (f.HasFlag(UserFlags.Special)) return Guest;
        return UserNameDefault;
    }

    /// <summary>Ported from bnubot's getChatColor(flags): colors a Talk line by the speaking user's highest-priority channel flag.</summary>
    public RgbColor GetChatColor(uint flags)
    {
        var f = (UserFlags)flags;
        if (f.HasFlag(UserFlags.Squelched)) return Whisper;
        if (f.HasFlag(UserFlags.Blizzard)) return Cyan;
        if (f.HasFlag(UserFlags.Admin)) return Green;
        if (f.HasFlag(UserFlags.Operator)) return White;
        if (f.HasFlag(UserFlags.Speaker)) return Speaker;
        if (f.HasFlag(UserFlags.Special)) return Guest;
        return White;
    }

    /// <summary>Ported from bnubot's getEmoteColor(flags): colors an Emote line by the speaking user's highest-priority channel flag.</summary>
    public RgbColor GetEmoteColor(uint flags)
    {
        var f = (UserFlags)flags;
        if (f.HasFlag(UserFlags.Blizzard)) return Cyan;
        if (f.HasFlag(UserFlags.Admin)) return Green;
        if (f.HasFlag(UserFlags.Operator)) return White;
        if (f.HasFlag(UserFlags.Speaker)) return Speaker;
        if (f.HasFlag(UserFlags.Special)) return Guest;
        return EmoteDefault;
    }

    /// <summary>
    /// This port's original palette — bnubot's InvigorationColorScheme, using
    /// the correctly byte-order-decoded values this port already shipped
    /// with (bnubot's own Java source reuses the raw VB6 BGR literals
    /// directly as RGB ints, a byte-order bug there — not reproduced here).
    /// </summary>
    public static readonly ChatPalette Invigoration = new()
    {
        Background = new RgbColor(0x24, 0x24, 0x24),
        White = new RgbColor(0xFF, 0xFF, 0xFF),
        Channel = new RgbColor(0x00, 0xCE, 0x00),
        Info = new RgbColor(0x00, 0xC0, 0xC0),
        Error = new RgbColor(0xCE, 0x3E, 0x3E),
        Debug = new RgbColor(0xCE, 0x88, 0x00),
        Gray = new RgbColor(0x55, 0x55, 0x55),
        SelfUserName = new RgbColor(0x2C, 0xAC, 0xE8),
        Whisper = new RgbColor(0xAF, 0xAF, 0xAF),
        Highlight = new RgbColor(0x8D, 0x00, 0xCE),
        Red = new RgbColor(0xCE, 0x3E, 0x3E),
        Green = new RgbColor(0x00, 0xCE, 0x00),
        Cyan = new RgbColor(0x00, 0xFF, 0xFF),
        Speaker = new RgbColor(0x51, 0xCE, 0xCE),
        Guest = new RgbColor(0x8D, 0x00, 0xCE),
        UserNameDefault = new RgbColor(0xA8, 0x9D, 0x65),
        EmoteDefault = new RgbColor(0xA8, 0x9D, 0x65),
    };

    /// <summary>bnubot's StarcraftColorScheme — plain java.awt.Color named constants, no byte-order ambiguity. Background is lightened from bnubot's pure black, which left pure-blue Info text nearly illegible.</summary>
    public static readonly ChatPalette StarCraft = new()
    {
        Background = new RgbColor(0x30, 0x30, 0x38),
        White = new RgbColor(0xFF, 0xFF, 0xFF), // Color.WHITE
        Channel = new RgbColor(0x00, 0xFF, 0x00), // Color.GREEN
        Info = new RgbColor(0x5C, 0x8D, 0xFF), // lightened from Color.BLUE (0000FF) for legibility
        Error = new RgbColor(0xFF, 0x00, 0x00), // Color.RED
        Debug = new RgbColor(0xFF, 0xFF, 0x00), // Color.YELLOW
        Gray = new RgbColor(0x80, 0x80, 0x80), // Color.GRAY
        SelfUserName = new RgbColor(0x00, 0xFF, 0xFF), // Color.CYAN
        Whisper = new RgbColor(0x80, 0x80, 0x80), // Color.GRAY
        Highlight = new RgbColor(0xFF, 0x00, 0xFF), // Color.MAGENTA
        Red = new RgbColor(0xFF, 0x00, 0x00),
        Green = new RgbColor(0x00, 0xFF, 0x00),
        Cyan = new RgbColor(0x00, 0xFF, 0xFF),
        Speaker = new RgbColor(0xFF, 0xFF, 0x00),
        Guest = new RgbColor(0xFF, 0x00, 0xFF),
        UserNameDefault = new RgbColor(0xFF, 0xFF, 0x00),
        EmoteDefault = new RgbColor(0xFF, 0xFF, 0x00),
    };

    /// <summary>bnubot's Diablo2ColorScheme — already standard RGB hex literals, no byte-order ambiguity.</summary>
    public static readonly ChatPalette DiabloII = new()
    {
        Background = new RgbColor(0x08, 0x08, 0x08),
        White = new RgbColor(0xD0, 0xD0, 0xD0),
        Channel = new RgbColor(0x00, 0xCE, 0x00),
        Info = new RgbColor(0x44, 0x40, 0x9C),
        Error = new RgbColor(0xCE, 0x3E, 0x3E),
        Debug = new RgbColor(0xCE, 0xCE, 0x51),
        Gray = new RgbColor(0x55, 0x55, 0x55),
        SelfUserName = new RgbColor(0x00, 0xD0, 0xD0),
        Whisper = new RgbColor(0x55, 0x55, 0x55),
        Highlight = new RgbColor(0x8D, 0x00, 0xCE),
        Red = new RgbColor(0xCE, 0x3E, 0x3E),
        Green = new RgbColor(0x00, 0xCE, 0x00),
        Cyan = new RgbColor(0x00, 0xD0, 0xD0),
        Speaker = new RgbColor(0xCE, 0xCE, 0x51),
        Guest = new RgbColor(0x8D, 0x00, 0xCE),
        UserNameDefault = new RgbColor(0xA8, 0x9D, 0x65),
        EmoteDefault = new RgbColor(0x55, 0x55, 0x55),
    };

    public static ChatPalette ForScheme(Config.BotConfig config) => config.ChatColorScheme switch
    {
        ChatColorScheme.StarCraft => StarCraft,
        ChatColorScheme.DiabloII => DiabloII,
        ChatColorScheme.Custom => FromCustom(config.CustomColors),
        _ => Invigoration,
    };

    public static ChatPalette FromCustom(Config.CustomChatPalette c) => new()
    {
        Background = FromPacked(c.Background),
        White = FromPacked(c.White),
        Channel = FromPacked(c.Channel),
        Info = FromPacked(c.Info),
        Error = FromPacked(c.Error),
        Debug = FromPacked(c.Debug),
        Gray = FromPacked(c.Gray),
        SelfUserName = FromPacked(c.SelfUserName),
        Whisper = FromPacked(c.Whisper),
        Highlight = FromPacked(c.Highlight),
        Red = FromPacked(c.Red),
        Green = FromPacked(c.Green),
        Cyan = FromPacked(c.Cyan),
        Speaker = FromPacked(c.Speaker),
        Guest = FromPacked(c.Guest),
        UserNameDefault = FromPacked(c.UserNameDefault),
        EmoteDefault = FromPacked(c.EmoteDefault),
    };

    private static RgbColor FromPacked(int rgb) => new((byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb);

    /// <summary>Packs this palette into a CustomChatPalette — used to seed the Colors library's built-in scheme files.</summary>
    public Config.CustomChatPalette ToCustom() => new()
    {
        Background = ToPacked(Background),
        White = ToPacked(White),
        Channel = ToPacked(Channel),
        Info = ToPacked(Info),
        Error = ToPacked(Error),
        Debug = ToPacked(Debug),
        Gray = ToPacked(Gray),
        SelfUserName = ToPacked(SelfUserName),
        Whisper = ToPacked(Whisper),
        Highlight = ToPacked(Highlight),
        Red = ToPacked(Red),
        Green = ToPacked(Green),
        Cyan = ToPacked(Cyan),
        Speaker = ToPacked(Speaker),
        Guest = ToPacked(Guest),
        UserNameDefault = ToPacked(UserNameDefault),
        EmoteDefault = ToPacked(EmoteDefault),
    };

    private static int ToPacked(RgbColor c) => (c.R << 16) | (c.G << 8) | c.B;
}
