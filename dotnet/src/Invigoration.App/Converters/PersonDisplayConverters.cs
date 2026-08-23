using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Stimpak;

namespace Invigoration.App.Converters;

/// <summary>
/// Stimpak's own Person.Name already bakes the clan tag in as a *prefix* — confirmed via the
/// Rust source (ChatUser::visible_name, native/superiority/core/src/games/sc2/chat/session.rs):
/// "&lt;{tag}&gt; {name}". This reorders it to a suffix ("Username &lt;TAG&gt;") instead, per
/// request — stripping the known prefix via the separate ClanTag property rather than
/// re-deriving formatting Stimpak already owns.
/// </summary>
public sealed class PersonNameWithTrailingClanTagConverter : IValueConverter
{
    public static readonly PersonNameWithTrailingClanTagConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is not Person person)
        {
            return "";
        }

        if (person.ClanTag is { Length: > 0 } tag)
        {
            var prefix = $"<{tag}> ";
            if (person.Name.StartsWith(prefix, StringComparison.Ordinal))
            {
                return $"{person.Name[prefix.Length..]} <{tag}>";
            }
        }

        return person.Name;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

/// <summary>A multi-line hover tooltip surfacing everything Stimpak's Person actually exposes — there isn't much beyond presence/clan/handle, but what's there is worth showing on demand rather than cluttering the row itself.</summary>
public sealed class PersonDetailsTooltipConverter : IValueConverter
{
    public static readonly PersonDetailsTooltipConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is not Person person)
        {
            return "";
        }

        var lines = new List<string> { person.Name };
        if (!string.IsNullOrEmpty(person.ClanTag))
        {
            lines.Add($"Clan: {person.ClanTag}");
        }

        lines.Add($"Status: {person.Presence}");
        lines.Add($"Handle: {person.Handle}");
        if (person.PresenceId is { } presenceId)
        {
            lines.Add($"Presence ID: {presenceId}");
        }

        return string.Join(Environment.NewLine, lines);
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}
