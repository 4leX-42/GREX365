using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace Grex365.App.Converters;

public sealed class BoolToBrushConverter : IValueConverter
{
    private static readonly Brush On = new SolidColorBrush(Color.FromRgb(0x22, 0xC5, 0x5E));
    private static readonly Brush Off = new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80));

    static BoolToBrushConverter()
    {
        On.Freeze();
        Off.Freeze();
    }

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is bool b && b ? On : Off;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

public sealed class BoolToOnOffConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is bool b && b ? "conectado" : "desconectado";

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
