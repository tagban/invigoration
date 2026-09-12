using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Layout;

namespace Invigoration.App.Converters;

/// <summary>
/// Picks one of two values by a bool, with both options given as the ConverterParameter in
/// "whenTrue|whenFalse" form (e.g. ConverterParameter="2|1" on a Grid.Row). Used by
/// BotTabView.axaml to reposition the Users/Friends/Clan panel between its normal right-hand
/// column and D2 Style's bottom dock (BotConfig.UseD2ChatLayout) without duplicating the whole
/// panel into two visual trees — one control, different Grid coordinates.
///
/// Unlike this folder's other converters (single-purpose, no parameter), this one is
/// parameterized: the alternative was five near-identical converters for Grid.Row, Grid.Column,
/// Grid.ColumnSpan, Grid.RowSpan and MaxHeight, all saying the same thing.
/// </summary>
public sealed class BoolToDoubleConverter : IValueConverter
{
    public static readonly BoolToDoubleConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (parameter is not string options)
        {
            throw new ArgumentException("ConverterParameter must be a \"whenTrue|whenFalse\" string.", nameof(parameter));
        }

        var separator = options.IndexOf('|');
        if (separator < 0)
        {
            throw new ArgumentException($"ConverterParameter \"{options}\" is missing its '|' separator.", nameof(parameter));
        }

        var chosen = value is true ? options[..separator] : options[(separator + 1)..];
        return double.Parse(chosen, CultureInfo.InvariantCulture);
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

/// <summary>
/// Same "whenTrue|whenFalse" idea as <see cref="BoolToDoubleConverter"/>, for a Thickness — e.g.
/// ConverterParameter="8,0,8,8|0,4,8,8". The Users/Friends/Clan panel wants different margins
/// docked at the bottom (full width, so it needs a left margin) than in its normal right-hand
/// column (where the GridSplitter already provides the left gap).
/// </summary>
public sealed class BoolToThicknessConverter : IValueConverter
{
    public static readonly BoolToThicknessConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (parameter is not string options)
        {
            throw new ArgumentException("ConverterParameter must be a \"whenTrue|whenFalse\" string.", nameof(parameter));
        }

        var separator = options.IndexOf('|');
        if (separator < 0)
        {
            throw new ArgumentException($"ConverterParameter \"{options}\" is missing its '|' separator.", nameof(parameter));
        }

        var chosen = value is true ? options[..separator] : options[(separator + 1)..];
        return Avalonia.Thickness.Parse(chosen);
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

/// <summary>
/// True → Horizontal, false → Vertical. Drives the Users list's VirtualizingStackPanel direction
/// so D2 Style's bottom dock flows user rows left-to-right as a portrait strip. Deliberately a
/// VirtualizingStackPanel in both directions rather than a plain StackPanel — this list routinely
/// holds thousands of rows under a mass-join flood, and losing virtualization here would undo the
/// load-testing work that made that survivable.
/// </summary>
public sealed class BoolToOrientationConverter : IValueConverter
{
    public static readonly BoolToOrientationConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value is true ? Orientation.Horizontal : Orientation.Vertical;

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}
