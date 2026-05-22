using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public sealed record PrivilegedRoleAssignment(
    string RoleName,
    string? RoleTemplateId,
    string MemberId,
    string? MemberDisplayName,
    string? MemberUpn,
    string? MemberType,
    string? MemberUserType,
    bool MemberAccountEnabled);

public sealed record PrivilegedRoleSummary(
    int TotalAssignments,
    int UniqueAdmins,
    int GlobalAdmins,
    int GuestsWithAdminRole,
    int DisabledWithAdminRole,
    int ServicePrincipalsWithAdminRole);

public static class PrivilegedRoleAuditAnalyzer
{
    public const string GlobalAdministratorTemplateId = "62e90394-69f5-4237-9190-012177145e10";

    private const int GlobalAdminMaxBeforeSprawl = 5;

    public static (PrivilegedRoleSummary Summary, IReadOnlyList<AuditFinding> Findings) Analyze(
        IEnumerable<PrivilegedRoleAssignment> assignments)
    {
        ArgumentNullException.ThrowIfNull(assignments);

        var findings = new List<AuditFinding>();
        var globalAdminMembers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var allAdminMembers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var seenGuestPairs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var seenDisabledPairs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var seenSpPairs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int totalAssignments = 0;
        int guestsCount = 0, disabledCount = 0, spCount = 0;

        foreach (var a in assignments)
        {
            if (string.IsNullOrWhiteSpace(a.MemberId) || string.IsNullOrWhiteSpace(a.RoleName))
            {
                continue;
            }
            totalAssignments++;
            allAdminMembers.Add(a.MemberId);

            var who = !string.IsNullOrWhiteSpace(a.MemberUpn)
                ? a.MemberUpn!
                : (a.MemberDisplayName ?? a.MemberId);

            var pairKey = $"{a.RoleName}|{a.MemberId}";

            if (string.Equals(a.RoleTemplateId, GlobalAdministratorTemplateId, StringComparison.OrdinalIgnoreCase))
            {
                globalAdminMembers.Add(a.MemberId);
            }

            var isSp = string.Equals(a.MemberType, "ServicePrincipal", StringComparison.OrdinalIgnoreCase);
            var isGuest = string.Equals(a.MemberUserType, "Guest", StringComparison.OrdinalIgnoreCase);

            if (isGuest && seenGuestPairs.Add(pairKey))
            {
                guestsCount++;
                findings.Add(new AuditFinding(
                    "Guest with admin role",
                    who,
                    $"Invitado con role '{a.RoleName}' — riesgo de privilege escalation.",
                    "ERROR"));
            }

            if (!a.MemberAccountEnabled && !isSp && seenDisabledPairs.Add(pairKey))
            {
                disabledCount++;
                findings.Add(new AuditFinding(
                    "Disabled account with admin role",
                    who,
                    $"Cuenta deshabilitada conserva role '{a.RoleName}'.",
                    "ERROR"));
            }

            if (isSp && seenSpPairs.Add(pairKey))
            {
                spCount++;
                findings.Add(new AuditFinding(
                    "Service principal with admin role",
                    who,
                    $"Service principal con role '{a.RoleName}' — verifica si es app de gestión legitima.",
                    "INFO"));
            }
        }

        var globalAdmins = globalAdminMembers.Count;
        if (globalAdmins == 0)
        {
            findings.Add(new AuditFinding(
                "No Global Administrators",
                "(tenant)",
                "El tenant no tiene Global Administrators — pérdida de control administrativo.",
                "ERROR"));
        }
        else if (globalAdmins == 1)
        {
            findings.Add(new AuditFinding(
                "Single Global Administrator",
                "(tenant)",
                "Solo 1 Global Administrator — sin cuenta backup en caso de lockout.",
                "WARN"));
        }
        else if (globalAdmins > GlobalAdminMaxBeforeSprawl)
        {
            findings.Add(new AuditFinding(
                "Too many Global Administrators",
                "(tenant)",
                $"{globalAdmins} Global Administrators (>{GlobalAdminMaxBeforeSprawl}) — superficie de ataque elevada.",
                "WARN"));
        }

        var summary = new PrivilegedRoleSummary(
            TotalAssignments: totalAssignments,
            UniqueAdmins: allAdminMembers.Count,
            GlobalAdmins: globalAdmins,
            GuestsWithAdminRole: guestsCount,
            DisabledWithAdminRole: disabledCount,
            ServicePrincipalsWithAdminRole: spCount);

        return (summary, findings);
    }
}
