using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public sealed record TransportRuleSnapshot(
    string Name,
    string State,
    int Priority,
    string Mode,
    string? Description,
    IReadOnlyList<string> ForwardTo,
    IReadOnlyList<string> BlindCopyTo,
    IReadOnlyList<string> RedirectMessageTo,
    string? RouteMessageOutboundConnector,
    bool DeleteMessage,
    string? SentToScope,
    string? FromScope);

public sealed record TransportRulesSummary(
    int Total,
    int Enabled,
    int Disabled,
    int WithExternalForward,
    int WithExternalBcc,
    int WithExternalRedirect);

public static class TransportRuleAuditAnalyzer
{
    private static readonly string[] SecurityRuleNameKeywords =
    {
        "anti", "spam", "phish", "phishing", "malware", "dlp", "quarantine",
        "block", "spoof", "security", "external sender", "warning"
    };

    public static (TransportRulesSummary Summary, IReadOnlyList<AuditFinding> Findings) Analyze(
        IEnumerable<TransportRuleSnapshot> rules,
        IEnumerable<string> acceptedDomains)
    {
        ArgumentNullException.ThrowIfNull(rules);
        ArgumentNullException.ThrowIfNull(acceptedDomains);

        var domainSet = acceptedDomains
            .Where(d => !string.IsNullOrWhiteSpace(d))
            .Select(d => d.Trim().TrimStart('@').ToLowerInvariant())
            .ToHashSet();

        var findings = new List<AuditFinding>();
        int total = 0, enabled = 0, disabled = 0;
        int extFwd = 0, extBcc = 0, extRedirect = 0;

        foreach (var r in rules)
        {
            if (string.IsNullOrWhiteSpace(r.Name))
            {
                continue;
            }
            total++;
            var isEnabled = string.Equals(r.State, "Enabled", StringComparison.OrdinalIgnoreCase);
            if (isEnabled)
            {
                enabled++;
            }
            else
            {
                disabled++;
                if (IsSecurityRuleByName(r.Name))
                {
                    findings.Add(new AuditFinding(
                        "Transport rule disabled (security keyword)",
                        r.Name,
                        $"Rule disabled (state={r.State}) cuyo nombre sugiere seguridad/antispam — verifica si fue intencional.",
                        "WARN"));
                }
                else
                {
                    findings.Add(new AuditFinding(
                        "Transport rule disabled",
                        r.Name,
                        $"Rule disabled (state={r.State}).",
                        "INFO"));
                }
                continue;
            }

            var externalForwards = ExternalRecipients(r.ForwardTo, domainSet);
            if (externalForwards.Count > 0)
            {
                extFwd++;
                findings.Add(new AuditFinding(
                    "Transport rule forwards externally",
                    r.Name,
                    $"Enabled rule reenvía a destinatarios externos: {string.Join(", ", externalForwards)}. Vector exfil.",
                    "ERROR"));
            }

            var externalBcc = ExternalRecipients(r.BlindCopyTo, domainSet);
            if (externalBcc.Count > 0)
            {
                extBcc++;
                findings.Add(new AuditFinding(
                    "Transport rule BCC externally",
                    r.Name,
                    $"Enabled rule BCC silencioso a destinatarios externos: {string.Join(", ", externalBcc)}. Vector exfil.",
                    "ERROR"));
            }

            var externalRedirect = ExternalRecipients(r.RedirectMessageTo, domainSet);
            if (externalRedirect.Count > 0)
            {
                extRedirect++;
                findings.Add(new AuditFinding(
                    "Transport rule redirects externally",
                    r.Name,
                    $"Enabled rule redirige a destinatarios externos: {string.Join(", ", externalRedirect)}. Vector exfil.",
                    "ERROR"));
            }

            if (!string.IsNullOrWhiteSpace(r.RouteMessageOutboundConnector))
            {
                findings.Add(new AuditFinding(
                    "Transport rule routes via custom connector",
                    r.Name,
                    $"Outbound connector custom='{r.RouteMessageOutboundConnector}' — verifica destino legítimo (escenarios hybrid normales).",
                    "INFO"));
            }

            if (r.DeleteMessage && IsBroadScope(r))
            {
                findings.Add(new AuditFinding(
                    "Transport rule deletes with broad scope",
                    r.Name,
                    "DeleteMessage activo sin condiciones específicas — riesgo de pérdida silenciosa de correo.",
                    "WARN"));
            }

            if (string.Equals(r.Mode, "Audit", StringComparison.OrdinalIgnoreCase))
            {
                findings.Add(new AuditFinding(
                    "Transport rule in audit mode",
                    r.Name,
                    "Mode=Audit (no enforcement) — rule no actúa, solo registra. ¿Falta promoverla?",
                    "INFO"));
            }
        }

        var summary = new TransportRulesSummary(
            Total: total,
            Enabled: enabled,
            Disabled: disabled,
            WithExternalForward: extFwd,
            WithExternalBcc: extBcc,
            WithExternalRedirect: extRedirect);

        return (summary, findings);
    }

    private static bool IsSecurityRuleByName(string name)
    {
        var lower = name.ToLowerInvariant();
        return SecurityRuleNameKeywords.Any(k => lower.Contains(k));
    }

    private static List<string> ExternalRecipients(IReadOnlyList<string> recipients, HashSet<string> acceptedDomains)
    {
        var external = new List<string>();
        foreach (var raw in recipients)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                continue;
            }
            var addr = raw.Trim();
            if (addr.StartsWith("smtp:", StringComparison.OrdinalIgnoreCase))
            {
                addr = addr[5..];
            }
            var at = addr.LastIndexOf('@');
            if (at <= 0 || at >= addr.Length - 1)
            {
                // No domain (likely internal mailbox alias) — skip, not external.
                continue;
            }
            var domain = addr[(at + 1)..].TrimEnd('>').TrimEnd('.').ToLowerInvariant();
            if (!acceptedDomains.Contains(domain))
            {
                external.Add(addr);
            }
        }
        return external;
    }

    private static bool IsBroadScope(TransportRuleSnapshot rule) =>
        string.IsNullOrWhiteSpace(rule.SentToScope) && string.IsNullOrWhiteSpace(rule.FromScope);
}
