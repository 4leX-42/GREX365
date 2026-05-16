namespace Grex365.Core.Models;

public sealed record GroupSummary(
    string Id,
    string DisplayName,
    string? Mail,
    string GroupKind);

public sealed record GroupMember(
    string Id,
    string? DisplayName,
    string? Mail,
    string? UserPrincipalName);

public sealed record AddMemberResult(
    string Input,
    string Status,
    string Detail);
