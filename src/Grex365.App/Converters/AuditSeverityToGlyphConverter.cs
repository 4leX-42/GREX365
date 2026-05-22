using System.Globalization;
using System.Windows.Data;

namespace Grex365.App.Converters;

public sealed class AuditSeverityToGlyphConverter : IValueConverter
{
    // Segoe Fluent Icons codepoints (escaped to survive editor transports).
    private const string ErrorGlyph = "";   // ErrorBadge
    private const string WarningGlyph = ""; // Warning
    private const string InfoGlyph = "";    // Info
    private const string DefaultGlyph = ""; // UnknownMirrored

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
        (value as string)?.ToUpperInvariant() switch
        {
            "ERROR" => ErrorGlyph,
            "WARN" => WarningGlyph,
            "INFO" => InfoGlyph,
            _ => DefaultGlyph,
        };

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
