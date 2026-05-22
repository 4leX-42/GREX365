using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace Grex365.App.Converters;

public sealed class AuditSeverityToBrushConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var key = (value as string)?.ToUpperInvariant() switch
        {
            "ERROR" => "BrushSemanticError",
            "WARN" => "BrushSemanticWarn",
            "INFO" => "BrushSemanticInfo",
            _ => "BrushSemanticNeutral",
        };
        return LookupBrush(key);
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();

    private static Brush LookupBrush(string key)
    {
        if (Application.Current?.TryFindResource(key) is Brush b) return b;
        // Fallback in case resources not loaded yet (designer / unit test).
        return new SolidColorBrush(Color.FromRgb(0x9C, 0xA3, 0xAF));
    }
}
