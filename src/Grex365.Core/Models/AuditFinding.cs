namespace Grex365.Core.Models;

public sealed record AuditFinding(
    string Category,
    string Identity,
    string Detail,
    string Severity);

public sealed record AuditSummary(
    int UsersTotal,
    int UsersEnabled,
    int UsersDisabled,
    int Guests,
    int StaleMembers,
    int StaleGuests,
    int DisabledWithLicense);
