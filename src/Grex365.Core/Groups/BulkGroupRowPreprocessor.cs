using Grex365.Core.Models;

namespace Grex365.Core.Groups;

public static class BulkGroupRowPreprocessor
{
    public static IReadOnlyList<BulkGroupRow> Normalize(IEnumerable<IReadOnlyDictionary<string, string>> rawRows)
    {
        var output = new List<BulkGroupRow>();
        string lastGroupName = string.Empty;
        string lastGroupType = "M365";

        foreach (var row in rawRows)
        {
            row.TryGetValue("GroupName", out var groupName);
            row.TryGetValue("Email", out var email);
            var groupType = ReadGroupType(row);

            groupName = (groupName ?? string.Empty).Trim();
            email = (email ?? string.Empty).Trim();

            if (!string.IsNullOrEmpty(groupName))
            {
                lastGroupName = groupName;
                lastGroupType = groupType ?? "M365";
            }
            else
            {
                groupName = lastGroupName;
                groupType ??= lastGroupType;
            }

            if (string.IsNullOrEmpty(groupName) || string.IsNullOrEmpty(email))
            {
                continue;
            }
            output.Add(new BulkGroupRow(groupName, email, groupType ?? "M365"));
        }

        return output;
    }

    private static string? ReadGroupType(IReadOnlyDictionary<string, string> row)
    {
        foreach (var col in new[] { "GroupType", "Type", "Kind" })
        {
            if (row.TryGetValue(col, out var raw) && !string.IsNullOrWhiteSpace(raw))
            {
                return NormalizeType(raw);
            }
        }
        return null;
    }

    public static string NormalizeType(string raw)
    {
        var trimmed = (raw ?? string.Empty).Trim().ToLowerInvariant();
        return trimmed switch
        {
            "m365" or "microsoft 365" or "unified" or "group" or "office365" => "M365",
            "dl" or "distribution" or "distributionlist" or "distribution list" or "exchange" => "DL",
            _ => "M365",
        };
    }

    public static bool IsEmail(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }
        var at = value.IndexOf('@');
        if (at <= 0 || at == value.Length - 1)
        {
            return false;
        }
        var dot = value.IndexOf('.', at);
        return dot > at + 1 && dot < value.Length - 1;
    }
}
