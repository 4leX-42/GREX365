using System.Globalization;
using System.Windows.Data;

namespace Grex365.App.Converters;

public sealed class InitialsConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var s = value as string;
        if (string.IsNullOrWhiteSpace(s)) return "?";
        var parts = s.Trim().Split(new[] { ' ', '.', '-', '_', '@' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0) return "?";
        if (parts.Length == 1) return char.ToUpperInvariant(parts[0][0]).ToString();
        return string.Concat(
            char.ToUpperInvariant(parts[0][0]),
            char.ToUpperInvariant(parts[1][0]));
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
