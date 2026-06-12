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

// Outcome of removing a member from one classic DL / mail-enabled security group via EXO
// (Remove-DistributionGroupMember). Group echoes the identity the cmdlet was invoked with.
public sealed record DistributionGroupRemovalResult(
    string Group,
    bool Success,
    string Detail);

// An ACTIVE Entra directory role the user holds (memberOf). PIM-eligible assignments don't
// surface here — only what is currently activated.
public sealed record DirectoryRoleSummary(
    string Id,
    string DisplayName);
