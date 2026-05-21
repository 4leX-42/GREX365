using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public sealed record MfaRegistrationRow(
    string UserPrincipalName,
    string? DisplayName,
    string? UserType,
    bool IsAdmin,
    bool IsMfaRegistered,
    bool IsMfaCapable);

public sealed record MfaCoverageSummary(
    int Total,
    int AdminsTotal,
    int AdminsWithoutMfa,
    int MembersTotal,
    int MembersWithoutMfa,
    int GuestsTotal,
    int GuestsWithoutMfa);

public static class MfaCoverageAnalyzer
{
    public static (MfaCoverageSummary Summary, IReadOnlyList<AuditFinding> Findings) Analyze(
        IEnumerable<MfaRegistrationRow> rows)
    {
        ArgumentNullException.ThrowIfNull(rows);

        var findings = new List<AuditFinding>();
        int total = 0, adminsTotal = 0, adminsMissing = 0;
        int membersTotal = 0, membersMissing = 0;
        int guestsTotal = 0, guestsMissing = 0;

        foreach (var row in rows)
        {
            if (string.IsNullOrWhiteSpace(row.UserPrincipalName))
            {
                continue;
            }
            total++;
            var isGuest = string.Equals(row.UserType, "Guest", StringComparison.OrdinalIgnoreCase);

            if (row.IsAdmin)
            {
                adminsTotal++;
                if (!row.IsMfaRegistered)
                {
                    adminsMissing++;
                    var capable = row.IsMfaCapable ? "capable=true" : "capable=false";
                    findings.Add(new AuditFinding(
                        "MFA missing (admin)",
                        row.UserPrincipalName,
                        $"Admin sin MFA registrado ({capable}) — riesgo crítico",
                        "ERROR"));
                }
            }
            else if (isGuest)
            {
                guestsTotal++;
                if (!row.IsMfaRegistered)
                {
                    guestsMissing++;
                    findings.Add(new AuditFinding(
                        "MFA missing (guest)",
                        row.UserPrincipalName,
                        "Invitado sin MFA registrado",
                        "INFO"));
                }
            }
            else
            {
                membersTotal++;
                if (!row.IsMfaRegistered)
                {
                    membersMissing++;
                    findings.Add(new AuditFinding(
                        "MFA missing (member)",
                        row.UserPrincipalName,
                        "Miembro sin MFA registrado",
                        "WARN"));
                }
            }
        }

        var summary = new MfaCoverageSummary(
            Total: total,
            AdminsTotal: adminsTotal,
            AdminsWithoutMfa: adminsMissing,
            MembersTotal: membersTotal,
            MembersWithoutMfa: membersMissing,
            GuestsTotal: guestsTotal,
            GuestsWithoutMfa: guestsMissing);

        return (summary, findings);
    }
}
