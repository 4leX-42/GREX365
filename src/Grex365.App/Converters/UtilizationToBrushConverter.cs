using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace Grex365.App.Converters;

public sealed class UtilizationToBrushConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var key = (value as string)?.ToUpperInvariant() switch
        {
            "LOW" => "BrushSemanticOk",
            "MEDIUM" => "BrushSemanticInfo",
            "HIGH" => "BrushSemanticWarn",
            "CRITICAL" => "BrushSemanticError",
            _ => "BrushSemanticNeutral",
        };
        if (Application.Current?.TryFindResource(key) is Brush b) return b;
        return new SolidColorBrush(Color.FromRgb(0x9C, 0xA3, 0xAF));
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
