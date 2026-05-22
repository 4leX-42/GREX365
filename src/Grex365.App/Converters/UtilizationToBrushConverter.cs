using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace Grex365.App.Converters;

public sealed class UtilizationToBrushConverter : IValueConverter
{
    private static readonly Brush LowBrush      = Freeze(new SolidColorBrush(Color.FromRgb(0x22, 0xC5, 0x5E))); // green
    private static readonly Brush MediumBrush   = Freeze(new SolidColorBrush(Color.FromRgb(0x3B, 0x82, 0xF6))); // blue
    private static readonly Brush HighBrush     = Freeze(new SolidColorBrush(Color.FromRgb(0xF5, 0x9E, 0x0B))); // amber
    private static readonly Brush CriticalBrush = Freeze(new SolidColorBrush(Color.FromRgb(0xEF, 0x44, 0x44))); // red
    private static readonly Brush DefaultBrush  = Freeze(new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80)));

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
        (value as string)?.ToUpperInvariant() switch
        {
            "LOW" => LowBrush,
            "MEDIUM" => MediumBrush,
            "HIGH" => HighBrush,
            "CRITICAL" => CriticalBrush,
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
