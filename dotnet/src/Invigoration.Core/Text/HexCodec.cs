using System.Text;

namespace Invigoration.Core.Text;

/// <summary>Port of StrToHex/HexToStr in modFunctions.bas — used by the "hex" chat command.</summary>
public static class HexCodec
{
    public static string StrToHex(string text)
    {
        var sb = new StringBuilder(text.Length * 2);
        foreach (var c in text)
        {
            sb.Append(((byte)c).ToString("X2"));
        }

        return sb.ToString();
    }

    public static string HexToStr(string hex)
    {
        hex = hex.Replace(" ", "");
        if (hex.Length % 2 != 0)
        {
            return "";
        }

        var sb = new StringBuilder(hex.Length / 2);
        for (var i = 0; i < hex.Length; i += 2)
        {
            sb.Append((char)Convert.ToInt32(hex.Substring(i, 2), 16));
        }

        return sb.ToString();
    }
}
