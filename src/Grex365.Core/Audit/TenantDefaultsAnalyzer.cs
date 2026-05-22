using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public sealed record AuthorizationPolicySnapshot(
    bool AllowedToSignUpEmailBasedSubscriptions,
    bool AllowedToUseSspr,
    bool AllowEmailVerifiedUsersToJoinOrganization,
    string? AllowInvitesFrom,
    bool DefaultUserCanCreateApps,
    bool DefaultUserCanCreateSecurityGroups,
    bool DefaultUserCanCreateTenants,
    bool DefaultUserCanReadOtherUsers);

public sealed record TenantDefaultsSummary(
    bool SecurityDefaultsEnabled,
    int Findings);

public static class TenantDefaultsAnalyzer
{
    public static (TenantDefaultsSummary Summary, IReadOnlyList<AuditFinding> Findings) Analyze(
        AuthorizationPolicySnapshot policy,
        bool securityDefaultsEnabled)
    {
        ArgumentNullException.ThrowIfNull(policy);
        var findings = new List<AuditFinding>();

        if (policy.DefaultUserCanCreateApps)
        {
            findings.Add(new AuditFinding(
                "Tenant default: users can create apps",
                "(tenant)",
                "defaultUserRolePermissions.allowedToCreateApps=true — usuarios pueden registrar apps (vector consent attacks).",
                "WARN"));
        }

        if (policy.DefaultUserCanCreateTenants)
        {
            findings.Add(new AuditFinding(
                "Tenant default: users can create tenants",
                "(tenant)",
                "defaultUserRolePermissions.allowedToCreateTenants=true — cualquier usuario puede crear nuevos tenants (rara vez intencional).",
                "WARN"));
        }

        if (policy.AllowEmailVerifiedUsersToJoinOrganization)
        {
            findings.Add(new AuditFinding(
                "Tenant default: email-verified self-join",
                "(tenant)",
                "allowedToSignUpEmailBasedSubscriptions/JoinOrganization=true — usuarios con email verificado se unen al tenant solos.",
                "WARN"));
        }

        if (string.Equals(policy.AllowInvitesFrom, "everyone", StringComparison.OrdinalIgnoreCase))
        {
            findings.Add(new AuditFinding(
                "Tenant default: anyone can invite guests",
                "(tenant)",
                "allowInvitesFrom=everyone — cualquier user/guest puede invitar más guests. Recomendado: adminsAndGuestInviters o adminsOnly.",
                "WARN"));
        }

        if (!policy.AllowedToUseSspr)
        {
            findings.Add(new AuditFinding(
                "Tenant default: SSPR disabled",
                "(tenant)",
                "allowedToUseSSPR=false — autoservicio de reset de password deshabilitado (incrementa carga helpdesk).",
                "INFO"));
        }

        if (policy.AllowedToSignUpEmailBasedSubscriptions)
        {
            findings.Add(new AuditFinding(
                "Tenant default: email-based subscriptions",
                "(tenant)",
                "allowedToSignUpEmailBasedSubscriptions=true — usuarios pueden self-suscribirse a servicios por email.",
                "INFO"));
        }

        if (securityDefaultsEnabled)
        {
            findings.Add(new AuditFinding(
                "Tenant: Security Defaults enabled",
                "(tenant)",
                "Security Defaults activos — MFA baseline gratis. Considera migrar a Conditional Access para granularidad.",
                "INFO"));
        }

        return (new TenantDefaultsSummary(securityDefaultsEnabled, findings.Count), findings);
    }
}
