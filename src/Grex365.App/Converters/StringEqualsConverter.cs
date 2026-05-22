using System.Globalization;
using System.Windows.Data;

namespace Grex365.App.Converters;

public sealed class StringEqualsConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return string.Equals(value as string, parameter as string, StringComparison.OrdinalIgnoreCase);
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is bool b && b && parameter is string p)
        {
            return p;
        }
        return Binding.DoNothing;
    }
}
