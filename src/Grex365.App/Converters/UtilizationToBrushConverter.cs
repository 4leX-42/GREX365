using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace Grex365.App.Converters;

public sealed class UtilizationToBrushConverter : IValueConverter
{
    // License utilization mapping: kills the green leak — "low utilization" is not
    // semantically "ok" (could mean wasted seats), so map it to neutral instead of
    // BrushSemanticOk. The amber/red high-end still signals capacity pressure.
    //   LOW (<60%)      → neutral slate (under-utilized, neither good nor bad)
    //   MEDIUM (60-79%) → brand accent blue (normal operating range)
    //   HIGH (80-94%)   → amber warn (approaching capacity)
    //   CRITICAL (≥95%) → red error (act now)
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var key = (value as string)?.ToUpperInvariant() switch
        {
            "LOW" => "BrushSemanticNeutral",
            "MEDIUM" => "BrandAccentSolid",
            "HIGH" => "BrushSemanticWarn",
            "CRITICAL" => "BrushSemanticError",
            _ => "BrushSemanticNeutral",
        };
        if (Application.Current?.TryFindResource(key) is Brush b) return b;
        return new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80));
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
