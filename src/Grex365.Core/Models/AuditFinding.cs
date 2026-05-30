namespace Grex365.Core.Models;

public sealed record AuditFinding(
    string Category,
    string Identity,
    string Detail,
    string Severity)
{
    // Category emitted for disabled accounts that still hold licenses — the one finding
    // the Audit view can auto-correct via the guided offboarding flow.
    public const string DisabledWithLicenseCategory = "Disabled+License";

    public bool IsAutoFixable =>
        string.Equals(Category, DisabledWithLicenseCategory, StringComparison.OrdinalIgnoreCase);
}

public sealed record AuditSummary(
    int UsersTotal,
    int UsersEnabled,
    int UsersDisabled,
    int Guests,
    int StaleMembers,
    int StaleGuests,
    int DisabledWithLicense);
