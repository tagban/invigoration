using System.Globalization;

namespace Invigoration.Core.Updates;

/// <summary>
/// One of this app's version strings — "2.0.8b", or "v2.0.8b" as a release tag carries it.
/// Numbers first, then the letter suffix that marks the beta line ("2.0.8b" is newer than
/// "2.0.8a", and "2.0.10b" is newer than "2.0.9b" — which a plain string comparison gets wrong,
/// hence parsing rather than comparing text).
/// </summary>
public sealed record ReleaseVersion(IReadOnlyList<int> Numbers, string Suffix) : IComparable<ReleaseVersion>
{
    /// <summary>Null when the text isn't a version at all — a tag naming something else, or a release title that got through.</summary>
    public static ReleaseVersion? TryParse(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        var span = text.Trim();
        if (span.StartsWith('v') || span.StartsWith('V'))
        {
            span = span[1..];
        }

        var suffixStart = span.Length;
        while (suffixStart > 0 && !char.IsAsciiDigit(span[suffixStart - 1]))
        {
            suffixStart--;
        }

        var suffix = span[suffixStart..];
        var numbers = new List<int>();
        foreach (var part in span[..suffixStart].Split('.', StringSplitOptions.RemoveEmptyEntries))
        {
            if (!int.TryParse(part, NumberStyles.None, CultureInfo.InvariantCulture, out var number))
            {
                return null;
            }

            numbers.Add(number);
        }

        return numbers.Count == 0 ? null : new ReleaseVersion(numbers, suffix);
    }

    public int CompareTo(ReleaseVersion? other)
    {
        if (other is null)
        {
            return 1;
        }

        for (var i = 0; i < Math.Max(Numbers.Count, other.Numbers.Count); i++)
        {
            // A missing part counts as 0, so "2.1" and "2.1.0" are the same version.
            var mine = i < Numbers.Count ? Numbers[i] : 0;
            var theirs = i < other.Numbers.Count ? other.Numbers[i] : 0;
            if (mine != theirs)
            {
                return mine.CompareTo(theirs);
            }
        }

        return string.CompareOrdinal(Suffix, other.Suffix);
    }

    public bool IsNewerThan(ReleaseVersion other) => CompareTo(other) > 0;

    public override string ToString() => string.Join('.', Numbers) + Suffix;
}
