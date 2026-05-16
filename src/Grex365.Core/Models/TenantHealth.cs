namespace Grex365.Core.Models;

public sealed record TenantHealth(
    string TenantId,
    string DisplayName,
    string? VerifiedDomain,
    int TotalUsers,
    int TotalGroups,
    IReadOnlyList<LicenseSummary> Licenses);

public sealed record LicenseSummary(
    string SkuPartNumber,
    string SkuId,
    int Consumed,
    int Enabled,
    int Warning,
    int Suspended);
