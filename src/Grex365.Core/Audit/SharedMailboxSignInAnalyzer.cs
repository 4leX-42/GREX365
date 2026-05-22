using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public sealed record SharedMailboxSignInRow(
    string UserPrincipalName,
    string? DisplayName,
    bool? AccountDisabled);    // null = unknown / lookup failed

public sealed record SharedMailboxSignInSummary(
    int Total,
    int SignInEnabled,
    int SignInDisabled,
    int Unknown);

public static class SharedMailboxSignInAnalyzer
{
    public static (SharedMailboxSignInSummary Summary, IReadOnlyList<AuditFinding> Findings) Analyze(
        IEnumerable<SharedMailboxSignInRow> rows)
    {
        ArgumentNullException.ThrowIfNull(rows);

        var findings = new List<AuditFinding>();
        int total = 0, enabled = 0, disabled = 0, unknown = 0;

        foreach (var r in rows)
        {
            if (string.IsNullOrWhiteSpace(r.UserPrincipalName))
            {
                continue;
            }
            total++;
            var who = r.UserPrincipalName;

            if (r.AccountDisabled == false)
            {
                enabled++;
                findings.Add(new AuditFinding(
                    "Shared mailbox sign-in enabled",
                    who,
                    "Shared mailbox con AccountDisabled=false — sign-in habilitado. Bloquéalo (Set-User -AccountDisabled $true) para evitar password attacks.",
                    "WARN"));
            }
            else if (r.AccountDisabled == true)
            {
                disabled++;
            }
            else
            {
                unknown++;
                findings.Add(new AuditFinding(
                    "Shared mailbox sign-in unknown",
                    who,
                    "No se pudo determinar AccountDisabled (Get-User devolvió null). Verifica permisos o existencia del recipient.",
                    "INFO"));
            }
        }

        var summary = new SharedMailboxSignInSummary(
            Total: total,
            SignInEnabled: enabled,
            SignInDisabled: disabled,
            Unknown: unknown);

        return (summary, findings);
    }
}
