using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace Grex365.App.Converters;

public sealed class AuditSeverityToBrushConverter : IValueConverter
{
    private static readonly Brush InfoBrush = Freeze(new SolidColorBrush(Color.FromRgb(0x3B, 0x82, 0xF6)));
    private static readonly Brush WarnBrush = Freeze(new SolidColorBrush(Color.FromRgb(0xF5, 0x9E, 0x0B)));
    private static readonly Brush ErrorBrush = Freeze(new SolidColorBrush(Color.FromRgb(0xEF, 0x44, 0x44)));
    private static readonly Brush DefaultBrush = Freeze(new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80)));

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
