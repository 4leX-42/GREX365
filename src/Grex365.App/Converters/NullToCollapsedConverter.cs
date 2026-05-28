using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Grex365.App.Converters;

public sealed class NullToCollapsedConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object parameter, CultureInfo culture)
    {
        var invert = string.Equals(parameter as string, "invert", StringComparison.OrdinalIgnoreCase);
        var hasValue = value is not null;
        if (invert) hasValue = !hasValue;
        return hasValue ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
