using Grex365.Core.Models;

namespace Grex365.Core.Groups;

public static class BulkGroupRowPreprocessor
{
    public static IReadOnlyList<BulkGroupRow> Normalize(IEnumerable<IReadOnlyDictionary<string, string>> rawRows)
    {
        var output = new List<BulkGroupRow>();
        string lastGroupName = string.Empty;

        foreach (var row in rawRows)
        {
            row.TryGetValue("GroupName", out var groupName);
            row.TryGetValue("Email", out var email);

            groupName = (groupName ?? string.Empty).Trim();
            email = (email ?? string.Empty).Trim();

            if (!string.IsNullOrEmpty(groupName))
            {
                lastGroupName = groupName;
            }
            else
            {
                groupName = lastGroupName;
            }

            if (string.IsNullOrEmpty(groupName) || string.IsNullOrEmpty(email))
            {
                continue;
            }
            output.Add(new BulkGroupRow(groupName, email));
        }

        return output;
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
