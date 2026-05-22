using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using Grex365.Core.Models;

namespace Grex365.App.Converters;

public sealed class SeverityToBrushConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var key = value switch
        {
            LogSeverity.Debug => "BrushSemanticDebug",
            LogSeverity.Info => "BrushSemanticInfo",
            LogSeverity.Ok => "BrushSemanticOk",
            LogSeverity.Warning => "BrushSemanticWarn",
            LogSeverity.Error => "BrushSemanticError",
            _ => "BrushSemanticNeutral",
        };
        if (Application.Current?.TryFindResource(key) is Brush b) return b;
        return new SolidColorBrush(Color.FromRgb(0x9C, 0xA3, 0xAF));
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
