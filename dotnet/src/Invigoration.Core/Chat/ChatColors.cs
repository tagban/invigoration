namespace Invigoration.Core.Chat;

/// <summary>Plain RGB triple. Kept UI-framework-agnostic so Invigoration.Core has no Avalonia dependency.</summary>
public readonly record struct RgbColor(byte R, byte G, byte B)
{
    /// <summary>
    /// VB6 color constants (and Win32 COLORREF) pack channels as 0x00BBGGRR,
    /// the reverse byte order from standard RGB hex literals.
    /// </summary>
    public static RgbColor FromWin32Bgr(uint bgr) => new(
        (byte)(bgr & 0xFF),
        (byte)((bgr >> 8) & 0xFF),
        (byte)((bgr >> 16) & 0xFF));
}

/// <summary>One colored run of text in a chat/log line; a line is a list of these.</summary>
public readonly record struct ChatLogSegment(RgbColor Color, string Text);

/// <summary>Named colors ported from the D2* constants in modDeclares.bas.</summary>
public static class ChatColors
{
    public static readonly RgbColor White = RgbColor.FromWin32Bgr(0xFFFFFF);
    public static readonly RgbColor Red = RgbColor.FromWin32Bgr(0x3E3ECE);
    public static readonly RgbColor Green = RgbColor.FromWin32Bgr(0x00CE00);
    public static readonly RgbColor Blue = RgbColor.FromWin32Bgr(0x9C4044);
    public static readonly RgbColor Gray = RgbColor.FromWin32Bgr(0x555555);
    public static readonly RgbColor Orange = RgbColor.FromWin32Bgr(0x0088CE);
    public static readonly RgbColor LtYellow = RgbColor.FromWin32Bgr(0x51CECE);
    public static readonly RgbColor Purple = RgbColor.FromWin32Bgr(0xCE008D);
    public static readonly RgbColor Cyan = RgbColor.FromWin32Bgr(0x00FFFF);
    public static readonly RgbColor MedBlue = RgbColor.FromWin32Bgr(0xE8AC2C);
    public static readonly RgbColor LtBlue = RgbColor.FromWin32Bgr(0xC0C000);
    public static readonly RgbColor HexPink = RgbColor.FromWin32Bgr(0x9900FF);
    public static readonly RgbColor Yellow = RgbColor.FromWin32Bgr(0x00FFFF);
}
