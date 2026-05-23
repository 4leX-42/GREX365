namespace Grex365.Core.Csv;

public static class CsvEscaper
{
    // Escape a single CSV field per RFC 4180:
    // - null/empty → empty string
    // - contains ',' / '"' / '\n' / '\r' → wrap in double-quotes + double internal quotes
    // - otherwise → returned as-is
    public static string Escape(string? value)
    {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        var needs = value.Contains(',')
            || value.Contains('"')
            || value.Contains('\n')
            || value.Contains('\r');
        return needs ? "\"" + value.Replace("\"", "\"\"") + "\"" : value;
    }
}
