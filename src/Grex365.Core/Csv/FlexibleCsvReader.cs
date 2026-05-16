using System.Text;

namespace Grex365.Core.Csv;

public static class FlexibleCsvReader
{
    public static IReadOnlyList<IReadOnlyDictionary<string, string>> Read(string path)
    {
        using var stream = File.OpenRead(path);
        return Read(stream);
    }

    public static IReadOnlyList<IReadOnlyDictionary<string, string>> Read(Stream stream)
    {
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        var firstLine = reader.ReadLine();
        if (string.IsNullOrEmpty(firstLine))
        {
            return Array.Empty<IReadOnlyDictionary<string, string>>();
        }

        var delimiter = DetectDelimiter(firstLine);
        var headers = ParseLine(firstLine, delimiter);
        var rows = new List<IReadOnlyDictionary<string, string>>();

        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            var fields = ParseLine(line, delimiter);
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < headers.Count; i++)
            {
                var key = headers[i].Trim();
                if (string.IsNullOrEmpty(key))
                {
                    continue;
                }
                dict[key] = i < fields.Count ? fields[i].Trim() : string.Empty;
            }
            rows.Add(dict);
        }
        return rows;
    }

    private static char DetectDelimiter(string header)
    {
        var semis = header.Count(c => c == ';');
        var commas = header.Count(c => c == ',');
        var tabs = header.Count(c => c == '\t');

        if (semis >= commas && semis >= tabs)
        {
            return ';';
        }
        if (tabs > commas)
        {
            return '\t';
        }
        return ',';
    }

    private static List<string> ParseLine(string line, char delimiter)
    {
        var fields = new List<string>();
        var sb = new StringBuilder();
        var inQuotes = false;

        for (var i = 0; i < line.Length; i++)
        {
            var c = line[i];
            if (inQuotes)
            {
                if (c == '"')
                {
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    sb.Append(c);
                }
            }
            else
            {
                if (c == '"')
                {
                    inQuotes = true;
                }
                else if (c == delimiter)
                {
                    fields.Add(sb.ToString());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }
        }
        fields.Add(sb.ToString());
        return fields;
    }
}
