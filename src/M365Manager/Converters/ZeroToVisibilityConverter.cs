using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace M365Manager.Converters;

/// <summary>
/// 0 -> Collapsed, anything else -> Visible. For hiding a list container when its collection is
/// empty, where the empty case is carried by something else (an AlertBox, an empty state).
/// </summary>
public sealed class ZeroToVisibilityConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is int count && count == 0 ? Visibility.Collapsed : Visibility.Visible;

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
