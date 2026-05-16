using Grex365.Core.Models;

namespace Grex365.Core.Users;

public enum BulkUserActionKind
{
    Unknown,
    Enable,
    Disable,
    RemoveLicenses,
    AssignLicense
}

public sealed record BulkUserAction(BulkUserActionKind Kind, string? SkuPartNumber);

public static class BulkUserActionParser
{
    public static BulkUserAction Parse(string? rawAction)
    {
        var s = (rawAction ?? string.Empty).Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(s))
        {
            return new BulkUserAction(BulkUserActionKind.Unknown, null);
        }
        return s switch
        {
            "enable" => new BulkUserAction(BulkUserActionKind.Enable, null),
            "disable" => new BulkUserAction(BulkUserActionKind.Disable, null),
            "remove-licenses" => new BulkUserAction(BulkUserActionKind.RemoveLicenses, null),
            _ when s.StartsWith("assign:", StringComparison.Ordinal) =>
                new BulkUserAction(BulkUserActionKind.AssignLicense, ExtractSku(s)),
            _ => new BulkUserAction(BulkUserActionKind.Unknown, null)
        };
    }

    public static SkuInfo? FindByPartNumber(IEnumerable<SkuInfo> available, string? partNumber)
    {
        if (string.IsNullOrWhiteSpace(partNumber))
        {
            return null;
        }
        return available.FirstOrDefault(s =>
            string.Equals(s.SkuPartNumber, partNumber, StringComparison.OrdinalIgnoreCase));
    }

    private static string? ExtractSku(string assignAction)
    {
        var colon = assignAction.IndexOf(':');
        if (colon < 0 || colon == assignAction.Length - 1)
        {
            return null;
        }
        var partRaw = assignAction[(colon + 1)..].Trim();
        return string.IsNullOrEmpty(partRaw) ? null : partRaw.ToUpperInvariant();
    }
}
