using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace Grex365.App.Converters;

public sealed class BoolToBrushConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var key = value is bool b && b ? "BrushSemanticOk" : "BrushSemanticNeutral";
        if (Application.Current?.TryFindResource(key) is Brush brush) return brush;
        return new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80));
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

public sealed class BoolToOnOffConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var b = value is bool bv && bv;
        if (parameter is string s && s.Contains('/'))
        {
            var parts = s.Split('/', 2);
            return b ? parts[0] : parts[1];
        }
        return b ? "conectado" : "desconectado";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
