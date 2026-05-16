namespace Grex365.Core.Models;

public sealed record BulkGroupRow(string GroupName, string Email);

public sealed record BulkGroupResult(
    string GroupName,
    string GroupEmail,
    string Action,
    string? UserEmail,
    string Detail);
