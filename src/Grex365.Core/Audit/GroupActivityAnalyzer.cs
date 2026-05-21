using System.Globalization;
using System.Text;
using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public sealed record GroupActivityRow(
    string DisplayName,
    bool IsDeleted,
    string? OwnerPrincipalName,
    DateOnly? LastActivityDate,
    string? GroupType,
    int MemberCount,
    int ExternalMemberCount);

public static class GroupActivityAnalyzer
{
    public static IReadOnlyList<GroupActivityRow> ParseCsv(Stream csvStream)
    {
        ArgumentNullException.ThrowIfNull(csvStream);

        using var reader = new StreamReader(csvStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        var headerLine = reader.ReadLine();
        if (string.IsNullOrEmpty(headerLine))
        {
            return Array.Empty<GroupActivityRow>();
        }

        var headers = SplitCsvLine(headerLine);
        var idx = BuildHeaderIndex(headers);
        var rows = new List<GroupActivityRow>();

        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }
            var fields = SplitCsvLine(line);
            rows.Add(new GroupActivityRow(
                DisplayName: Get(fields, idx, "Group Display Name"),
                IsDeleted: ParseBool(Get(fields, idx, "Is Deleted")),
                OwnerPrincipalName: NullIfEmpty(Get(fields, idx, "Owner Principal Name")),
                LastActivityDate: ParseDate(Get(fields, idx, "Last Activity Date")),
                GroupType: NullIfEmpty(Get(fields, idx, "Group Type")),
                MemberCount: ParseInt(Get(fields, idx, "Member Count")),
                ExternalMemberCount: ParseInt(Get(fields, idx, "External Member Count"))));
        }

        return rows;
    }

    public static IReadOnlyList<AuditFinding> Analyze(
        IEnumerable<GroupActivityRow> rows,
        DateOnly today,
        int inactivityDays)
    {
        ArgumentNullException.ThrowIfNull(rows);
        if (inactivityDays < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(inactivityDays), "Debe ser >= 1.");
        }

        var findings = new List<AuditFinding>();
        var cutoff = today.AddDays(-inactivityDays);

        foreach (var r in rows)
        {
            if (r.IsDeleted)
            {
                continue;
            }
            if (string.IsNullOrWhiteSpace(r.DisplayName))
            {
                continue;
            }

            var name = r.DisplayName;
            if (r.LastActivityDate is null)
            {
                findings.Add(new AuditFinding(
                    "Inactive M365 group",
                    name,
                    $"sin actividad registrada; members={r.MemberCount}",
                    "WARN"));
                continue;
            }

            if (r.LastActivityDate.Value < cutoff)
            {
                var last = r.LastActivityDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                findings.Add(new AuditFinding(
                    "Inactive M365 group",
                    name,
                    $"última actividad: {last} (>{inactivityDays}d); members={r.MemberCount}",
                    "WARN"));
            }
        }

        return findings;
    }

    private static Dictionary<string, int> BuildHeaderIndex(IReadOnlyList<string> headers)
    {
        var idx = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < headers.Count; i++)
        {
            var key = headers[i].Trim();
            if (string.IsNullOrEmpty(key))
            {
                continue;
            }
            idx[key] = i;
        }
        return idx;
    }

    private static string Get(IReadOnlyList<string> fields, Dictionary<string, int> idx, string column)
    {
        return idx.TryGetValue(column, out var i) && i < fields.Count
            ? fields[i].Trim()
            : string.Empty;
    }

    private static string? NullIfEmpty(string s) => string.IsNullOrWhiteSpace(s) ? null : s;

    private static bool ParseBool(string s)
        => bool.TryParse(s, out var b) && b;

    private static int ParseInt(string s)
        => int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0;

    private static DateOnly? ParseDate(string s)
    {
        if (string.IsNullOrWhiteSpace(s))
        {
            return null;
        }
        if (DateOnly.TryParseExact(s, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var d))
        {
            return d;
        }
        if (DateOnly.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out d))
        {
            return d;
        }
        return null;
    }

    private static List<string> SplitCsvLine(string line)
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
                else if (c == ',')
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
