using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Invigoration.App.Models;
using Invigoration.App.Views;
using Invigoration.Core.Sc2;
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

/// <summary>A StarCraft: Remastered member's game icon (see NativeMemberProducts), a StarCraft II member's portrait (NativeMemberPortraits), or null.</summary>
/// <remarks>Also a multi-value converter: the second value is only there to make the row redraw when icons change (BotTabViewModel.IconVersion).</remarks>
public sealed class PersonProductIconConverter : IValueConverter, IMultiValueConverter
{
    public static readonly PersonProductIconConverter Instance = new();

    public object? Convert(IList<object?> values, Type targetType, object? parameter, CultureInfo culture) =>
        Convert(values.Count > 0 ? values[0] : null, targetType, parameter, culture);

    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is not Person person)
        {
            return null;
        }

        // SC:R members: their classic game icon. SC2 members: their profile portrait, once downloaded.
        var name = BareName(person, culture);
        if (NativeMemberProducts.IconKeyFor(name) is { } key)
        {
            return GameIconLoader.Get(key);
        }

        return NativeMemberPortraits.For(name) is { } portrait ? Sc2PortraitImages.Get(portrait.Sheet, portrait.Cell) : null;
    }

    /// <summary>The member's name without the clan tag Stimpak's Person.Name can carry.</summary>
    internal static string BareName(Person person, CultureInfo culture)
    {
        var name = PersonNameWithTrailingClanTagConverter.Instance.Convert(person, typeof(string), null, culture) as string ?? person.Name;
        if (person.ClanTag is { Length: > 0 } tag && name.EndsWith($" <{tag}>", StringComparison.Ordinal))
        {
            name = name[..^(tag.Length + 3)];
        }

        return name;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

/// <summary>Whether to show a Person's presence dot: not for StarCraft: Remastered members, whose chat has no presence (it would always be "available").</summary>
public sealed class PersonShowsPresenceConverter : IValueConverter
{
    public static readonly PersonShowsPresenceConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value is not Person person || !NativeMemberProducts.IsKnown(PersonProductIconConverter.BareName(person, culture));

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

/// <summary>The light detail line under an SC2 member's name in the Full user list (NativeMemberPortraits.DetailFor). The second value only makes it refresh (BotTabViewModel.IconVersion).</summary>
public sealed class PersonDetailConverter : IMultiValueConverter
{
    public static readonly PersonDetailConverter Instance = new();

    public object? Convert(IList<object?> values, Type targetType, object? parameter, CultureInfo culture) =>
        values.Count > 0 && values[0] is Person person ? NativeMemberPortraits.DetailFor(PersonProductIconConverter.BareName(person, culture)) : "";
}

/// <summary>An SC2 member's dot: their status from presence (NativeMemberPortraits.StateFor) when known, else Stimpak's. The second value only makes it refresh.</summary>
public sealed class PersonPresenceStateConverter : IMultiValueConverter
{
    public static readonly PersonPresenceStateConverter Instance = new();

    public object? Convert(IList<object?> values, Type targetType, object? parameter, CultureInfo culture)
    {
        if (values.Count == 0 || values[0] is not Person person)
        {
            return PresenceState.Offline;
        }

        var state = NativeMemberPortraits.StateFor(PersonProductIconConverter.BareName(person, culture)) ?? person.Presence;
        return PresenceConverter.Instance.Convert(state, targetType, parameter, culture);
    }
}

/// <summary>
/// A name's main part and its "#1234" code, for showing the code dimmed: parameter "code" gives the
/// code (with the '#'), anything else the rest. Works on a Person or a plain string.
/// </summary>
public sealed class NameCodeConverter : IValueConverter
{
    public static readonly NameCodeConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var name = value is Person person
            ? PersonNameWithTrailingClanTagConverter.Instance.Convert(person, typeof(string), null, culture) as string ?? ""
            : value as string ?? "";
        var (main, code) = Invigoration.Core.Chat.NameParts.Split(name);
        return parameter as string == "code" ? code : main;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}
