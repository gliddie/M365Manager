using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace M365Manager.Converters;

/// <summary>true -> Visible, false -> Collapsed; with <see cref="Invert"/> the other way round.</summary>
public sealed class BoolToVisibilityConverter : IValueConverter
{
    /// <summary>Set on the resource instance to show the element when the bound value is false.</summary>
    public bool Invert { get; set; }

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is true != Invert ? Visibility.Visible : Visibility.Collapsed;

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is Visibility.Visible != Invert;
}
