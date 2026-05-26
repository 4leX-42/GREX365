using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace Grex365.App.Converters;

public sealed class BoolToBrushConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var key = value is bool b && b ? "BrandAccentSolid" : "BrushSemanticNeutral";
        if (Application.Current?.TryFindResource(key) is Brush brush) return brush;
        return new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80));
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

public sealed class BoolToOnOffConverter : IValueConverter
{
    // ConverterParameter conventions:
    //   - null/empty: defaults to Common.Connected / Common.Disconnected via L10n
    //   - "On/Off" (no dot): plain literal — backwards compatible with ad-hoc captions
    //   - "Key.On/Key.Off" (each part contains a dot): both parts treated as L10n keys
    //     and resolved at convert time so language switches without rebinding.
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var b = value is bool bv && bv;
        if (parameter is string s && s.Contains('/'))
        {
            var parts = s.Split('/', 2);
            return b ? Resolve(parts[0]) : Resolve(parts[1]);
        }
        return L10n.Get(b ? "Common.Connected" : "Common.Disconnected");
    }

    private static string Resolve(string token)
    {
        // Heuristic: dotted tokens are L10n keys (e.g. "Status.Enabled"). Plain
        // tokens stay literal so existing ad-hoc captions ("válido/inválido") work
        // until they migrate.
        if (token.Contains('.'))
        {
            return L10n.Get(token);
        }
        return token;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
