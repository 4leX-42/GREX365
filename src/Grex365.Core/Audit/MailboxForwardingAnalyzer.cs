using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public sealed record MailboxForwardingRow(
    string UserPrincipalName,
    string? ForwardingSmtpAddress,
    string? ForwardingAddress);

public static class MailboxForwardingAnalyzer
{
    public static IReadOnlyList<AuditFinding> Analyze(
        IEnumerable<MailboxForwardingRow> rows,
        IEnumerable<string> acceptedDomains)
    {
        ArgumentNullException.ThrowIfNull(rows);
        ArgumentNullException.ThrowIfNull(acceptedDomains);

        var accepted = new HashSet<string>(
            acceptedDomains.Where(d => !string.IsNullOrWhiteSpace(d)).Select(NormalizeDomain),
            StringComparer.OrdinalIgnoreCase);

        var findings = new List<AuditFinding>();

        foreach (var row in rows)
        {
            if (string.IsNullOrWhiteSpace(row.UserPrincipalName))
            {
                continue;
            }

            var smtp = ExtractSmtp(row.ForwardingSmtpAddress);
            if (smtp is not null)
            {
                var domain = DomainOf(smtp);
                if (domain is not null && !accepted.Contains(domain))
                {
                    findings.Add(new AuditFinding(
                        "External forwarding (SMTP)",
                        row.UserPrincipalName,
                        $"ForwardingSmtpAddress → {smtp} (dominio {domain} fuera del tenant)",
                        "WARN"));
                }
            }
        }

        return findings;
    }

    private static string? ExtractSmtp(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }
        // Exchange devuelve a veces "SMTP:user@domain"; quitar prefijo.
        var trimmed = raw.Trim();
        var colonIdx = trimmed.IndexOf(':');
        if (colonIdx >= 0
            && colonIdx <= 5
            && trimmed[..colonIdx].Equals("SMTP", StringComparison.OrdinalIgnoreCase))
        {
            trimmed = trimmed[(colonIdx + 1)..].Trim();
        }
        return string.IsNullOrWhiteSpace(trimmed) ? null : trimmed;
    }

    private static string? DomainOf(string smtp)
    {
        var at = smtp.LastIndexOf('@');
        if (at < 0 || at == smtp.Length - 1)
        {
            return null;
        }
        return NormalizeDomain(smtp[(at + 1)..]);
    }

    private static string NormalizeDomain(string domain)
        => domain.Trim().TrimEnd('.').ToLowerInvariant();
}
