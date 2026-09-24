using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace M365Manager.Converters;

/// <summary>
/// null -> Collapsed, non-null -> Visible. Set <see cref="Invert"/> for the opposite, which is how
/// an "empty state" panel is shown exactly when its counterpart is hidden.
/// </summary>
public sealed class NullToVisibilityConverter : IValueConverter
{
    public bool Invert { get; set; }

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var isNull = value is null;
        if (Invert)
            isNull = !isNull;
        return isNull ? Visibility.Collapsed : Visibility.Visible;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
