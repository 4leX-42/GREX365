using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public sealed record InboxRuleRow(
    string MailboxUpn,
    string RuleName,
    bool Enabled,
    bool DeleteMessage,
    string? MoveToFolder,
    IReadOnlyList<string> ForwardTo,
    IReadOnlyList<string> ForwardAsAttachmentTo,
    IReadOnlyList<string> RedirectTo,
    IReadOnlyList<string> SubjectContainsWords,
    IReadOnlyList<string> BodyContainsWords);

public static class InboxRuleAnalyzer
{
    // Folder names whose presence in MoveToFolder hides incoming mail from the user — typical BEC pattern.
    private static readonly HashSet<string> SuspiciousMoveTargets = new(StringComparer.OrdinalIgnoreCase)
    {
        "Deleted Items", "DeletedItems",
        "Junk Email", "Junk E-mail", "JunkEmail",
        "RSS Feeds", "RSS Subscriptions",
        "Archive",
        "Conversation History",
        "Notes",
    };

    // Words attackers commonly filter on to hide replies during invoice/wire fraud.
    private static readonly HashSet<string> SuspiciousKeywords = new(StringComparer.OrdinalIgnoreCase)
    {
        "invoice", "factura",
        "payment", "pago",
        "wire", "transferencia",
        "bank", "banco",
        "account", "cuenta",
        "password", "contraseña", "contrasena",
        "security", "seguridad",
        "fraud", "fraude",
        "scam", "estafa",
        "phish", "phishing",
        "suspicious", "sospechoso",
        "hack", "hacked",
    };

    public static IReadOnlyList<AuditFinding> Analyze(
        IEnumerable<InboxRuleRow> rules,
        IEnumerable<string> acceptedDomains)
    {
        ArgumentNullException.ThrowIfNull(rules);
        ArgumentNullException.ThrowIfNull(acceptedDomains);

        var accepted = new HashSet<string>(
            acceptedDomains.Where(d => !string.IsNullOrWhiteSpace(d)).Select(NormalizeDomain),
            StringComparer.OrdinalIgnoreCase);

        var findings = new List<AuditFinding>();

        foreach (var rule in rules)
        {
            if (string.IsNullOrWhiteSpace(rule.MailboxUpn))
            {
                continue;
            }
            if (!rule.Enabled)
            {
                continue;
            }

            var who = $"{rule.MailboxUpn} ← '{rule.RuleName}'";
            var keywordHit = MatchesSuspiciousKeyword(rule);

            if (rule.DeleteMessage)
            {
                var detail = keywordHit is not null
                    ? $"DeleteMessage=true filtrando '{keywordHit}' — clásico ocultamiento BEC"
                    : "DeleteMessage=true";
                var severity = keywordHit is not null ? "WARN" : "INFO";
                findings.Add(new AuditFinding("Inbox rule: delete", who, detail, severity));
            }

            if (!string.IsNullOrWhiteSpace(rule.MoveToFolder)
                && SuspiciousMoveTargets.Contains(rule.MoveToFolder!.Trim()))
            {
                var detail = keywordHit is not null
                    ? $"MoveToFolder='{rule.MoveToFolder}' filtrando '{keywordHit}' — ocultamiento BEC"
                    : $"MoveToFolder='{rule.MoveToFolder}'";
                var severity = keywordHit is not null ? "WARN" : "INFO";
                findings.Add(new AuditFinding("Inbox rule: hide", who, detail, severity));
            }

            var externalTargets = CollectExternalRecipients(rule, accepted);
            if (externalTargets.Count > 0)
            {
                findings.Add(new AuditFinding(
                    "Inbox rule: external forward",
                    who,
                    $"Forward/Redirect a {string.Join(", ", externalTargets)} (dominios externos)",
                    "WARN"));
            }
        }

        return findings;
    }

    private static string? MatchesSuspiciousKeyword(InboxRuleRow rule)
    {
        foreach (var w in rule.SubjectContainsWords.Concat(rule.BodyContainsWords))
        {
            if (string.IsNullOrWhiteSpace(w))
            {
                continue;
            }
            foreach (var token in w.Split(new[] { ' ', ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (SuspiciousKeywords.Contains(token.Trim()))
                {
                    return token.Trim();
                }
            }
        }
        return null;
    }

    private static List<string> CollectExternalRecipients(InboxRuleRow rule, HashSet<string> accepted)
    {
        var external = new List<string>();
        foreach (var target in rule.ForwardTo.Concat(rule.ForwardAsAttachmentTo).Concat(rule.RedirectTo))
        {
            var addr = ExtractSmtp(target);
            if (addr is null)
            {
                continue;
            }
            var domain = DomainOf(addr);
            if (domain is not null && !accepted.Contains(domain))
            {
                external.Add(addr);
            }
        }
        return external;
    }

    private static string? ExtractSmtp(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }
        var s = raw.Trim();
        // EXO usa formato "DisplayName [SMTP:user@domain]" o "user@domain".
        var smtpIdx = s.IndexOf("SMTP:", StringComparison.OrdinalIgnoreCase);
        if (smtpIdx >= 0)
        {
            var rest = s[(smtpIdx + 5)..];
            var closeIdx = rest.IndexOf(']');
            if (closeIdx > 0)
            {
                rest = rest[..closeIdx];
            }
            s = rest.Trim();
        }
        return s.Contains('@') ? s : null;
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
