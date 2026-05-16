namespace Grex365.Core.Models;

public sealed record BulkUserResult(
    string Upn,
    string Action,
    string Status,
    string Detail);

public sealed record UserSummary(
    string Id,
    string? DisplayName,
    string? UserPrincipalName,
    string? Mail,
    bool AccountEnabled,
    bool IsGuest,
    int AssignedLicenseCount,
    DateTimeOffset? LastSignIn);
