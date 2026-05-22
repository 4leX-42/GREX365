using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace Grex365.App.Converters;

public sealed class AuditSeverityToBrushConverter : IValueConverter
{
    // Brighter shades para legibilidad en dark theme (también funcionan en light).
    private static readonly Brush InfoBrush = Freeze(new SolidColorBrush(Color.FromRgb(0x60, 0xA5, 0xFA)));
    private static readonly Brush WarnBrush = Freeze(new SolidColorBrush(Color.FromRgb(0xFB, 0xBF, 0x24)));
    private static readonly Brush ErrorBrush = Freeze(new SolidColorBrush(Color.FromRgb(0xF8, 0x71, 0x71)));
    private static readonly Brush DefaultBrush = Freeze(new SolidColorBrush(Color.FromRgb(0x9C, 0xA3, 0xAF)));

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
        (value as string)?.ToUpperInvariant() switch
        {
            "ERROR" => ErrorBrush,
            "WARN" => WarnBrush,
            "INFO" => InfoBrush,
            _ => DefaultBrush,
        };

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();

    private static Brush Freeze(SolidColorBrush b)
    {
        b.Freeze();
        return b;
    }
}
