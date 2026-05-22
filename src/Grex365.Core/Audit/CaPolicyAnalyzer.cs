using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public sealed record CaPolicySnapshot(
    string Id,
    string DisplayName,
    string State,
    DateTimeOffset? CreatedDateTime,
    DateTimeOffset? ModifiedDateTime,
    IReadOnlyList<string> IncludeUsers,
    IReadOnlyList<string> ExcludeUsers,
    IReadOnlyList<string> IncludeGroups,
    IReadOnlyList<string> ExcludeGroups,
    IReadOnlyList<string> IncludeRoles,
    IReadOnlyList<string> ExcludeRoles,
    IReadOnlyList<string> IncludeApplications,
    IReadOnlyList<string> BuiltInControls,
    string? GrantOperator);

public sealed record CaPoliciesSummary(
    int Total,
    int Enabled,
    int Disabled,
    int ReportOnly,
    int WithoutEffectiveControls);

public static class CaPolicyAnalyzer
{
    private const string StateEnabled = "enabled";
    private const string StateDisabled = "disabled";
    private const string StateReportOnly = "enabledForReportingButNotEnforced";

    private const int ReportOnlyStaleDays = 30;

    public static (CaPoliciesSummary Summary, IReadOnlyList<AuditFinding> Findings) Analyze(
        IEnumerable<CaPolicySnapshot> policies,
        DateTimeOffset now)
    {
        ArgumentNullException.ThrowIfNull(policies);

        var findings = new List<AuditFinding>();
        int total = 0, enabled = 0, disabled = 0, reportOnly = 0, withoutControls = 0;

        foreach (var p in policies)
        {
            if (string.IsNullOrWhiteSpace(p.DisplayName))
            {
                continue;
            }
            total++;

            var state = p.State ?? string.Empty;
            var name = p.DisplayName;

            if (string.Equals(state, StateDisabled, StringComparison.OrdinalIgnoreCase))
            {
                disabled++;
                findings.Add(new AuditFinding(
                    "CA policy disabled",
                    name,
                    $"Policy desactivada (state={state}). Verifica si es intencional.",
                    "INFO"));
                continue;
            }

            if (string.Equals(state, StateReportOnly, StringComparison.OrdinalIgnoreCase))
            {
                reportOnly++;
                var modified = p.ModifiedDateTime ?? p.CreatedDateTime;
                if (modified.HasValue)
                {
                    var ageDays = (int)Math.Floor((now - modified.Value).TotalDays);
                    if (ageDays >= ReportOnlyStaleDays)
                    {
                        findings.Add(new AuditFinding(
                            "CA policy stale report-only",
                            name,
                            $"En report-only desde hace {ageDays}d (>= {ReportOnlyStaleDays}d). ¿Falta promoverla a Enabled?",
                            "WARN"));
                        continue;
                    }
                }
                findings.Add(new AuditFinding(
                    "CA policy report-only",
                    name,
                    "En report-only — recopilando señales, aún sin enforcement.",
                    "INFO"));
                continue;
            }

            if (!string.Equals(state, StateEnabled, StringComparison.OrdinalIgnoreCase))
            {
                findings.Add(new AuditFinding(
                    "CA policy unknown state",
                    name,
                    $"State desconocido: '{state}'.",
                    "WARN"));
                continue;
            }

            enabled++;

            if (p.BuiltInControls is null || p.BuiltInControls.Count == 0)
            {
                withoutControls++;
                findings.Add(new AuditFinding(
                    "CA policy without controls",
                    name,
                    "Policy enabled sin builtInControls — no aplica enforcement.",
                    "ERROR"));
                continue;
            }

            var controls = p.BuiltInControls.Select(c => c.Trim().ToLowerInvariant()).ToHashSet();
            var hasMfa = controls.Contains("mfa");
            var hasCompliant = controls.Contains("compliantdevice");
            var hasDomainJoined = controls.Contains("domainjoineddevice");
            var hasBlock = controls.Contains("block");

            if (!hasMfa && !hasCompliant && !hasDomainJoined && !hasBlock)
            {
                findings.Add(new AuditFinding(
                    "CA policy weak controls",
                    name,
                    $"Enabled pero sin MFA/compliantDevice/domainJoinedDevice/block. Controls=[{string.Join(",", p.BuiltInControls)}]",
                    "WARN"));
            }

            var includesAll = p.IncludeUsers.Any(u => string.Equals(u, "All", StringComparison.OrdinalIgnoreCase));
            var hasExclusions =
                p.ExcludeUsers.Count > 0 ||
                p.ExcludeGroups.Count > 0 ||
                p.ExcludeRoles.Count > 0;
            if (includesAll && !hasExclusions)
            {
                findings.Add(new AuditFinding(
                    "CA policy targets All without exclusions",
                    name,
                    "includeUsers=All sin exclusiones (riesgo lockout — añade break-glass account).",
                    "WARN"));
            }

            var hasUserScope =
                p.IncludeUsers.Count > 0 ||
                p.IncludeGroups.Count > 0 ||
                p.IncludeRoles.Count > 0;
            if (!hasUserScope)
            {
                findings.Add(new AuditFinding(
                    "CA policy without user scope",
                    name,
                    "Sin includeUsers/Groups/Roles — la policy no afecta a nadie.",
                    "ERROR"));
            }
        }

        if (total == 0)
        {
            findings.Add(new AuditFinding(
                "No CA policies",
                "(tenant)",
                "El tenant no tiene Conditional Access policies configuradas.",
                "ERROR"));
        }

        var summary = new CaPoliciesSummary(
            Total: total,
            Enabled: enabled,
            Disabled: disabled,
            ReportOnly: reportOnly,
            WithoutEffectiveControls: withoutControls);

        return (summary, findings);
    }
}
