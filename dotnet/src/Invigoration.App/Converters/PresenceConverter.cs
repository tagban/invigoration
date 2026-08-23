using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Invigoration.App.Models;
using Stimpak;

namespace Invigoration.App.Converters;

/// <summary>Maps Stimpak's own Presence enum (SC2/SC:R/WC3:R roster entries) onto the shared PresenceState the icon set understands, so PresenceIcon can bind directly to a Stimpak Person.</summary>
public sealed class PresenceConverter : IValueConverter
{
    public static readonly PresenceConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) => value switch
    {
        Presence.Available => PresenceState.Available,
        Presence.Away => PresenceState.Away,
        Presence.Busy => PresenceState.DoNotDisturb,
        Presence.InGame => PresenceState.InGame,
        _ => PresenceState.Offline,
    };

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}
