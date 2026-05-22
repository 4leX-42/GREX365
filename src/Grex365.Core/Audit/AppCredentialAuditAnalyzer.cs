using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public sealed record AppCredentialSnapshot(
    string AppId,
    string DisplayName,
    string CredentialType,    // "Password" | "Key"
    string? KeyId,
    string? CredentialDisplayName,
    DateTimeOffset? EndDateTime);

public sealed record AppCredentialsSummary(
    int Total,
    int Expired,
    int ExpiringSoon,
    int LongLived);

public static class AppCredentialAuditAnalyzer
{
    public const int ExpiringSoonDays = 30;
    public const int LongLivedDays = 730;

    public static (AppCredentialsSummary Summary, IReadOnlyList<AuditFinding> Findings) Analyze(
        IEnumerable<AppCredentialSnapshot> credentials,
        DateTimeOffset now)
    {
        ArgumentNullException.ThrowIfNull(credentials);

        var findings = new List<AuditFinding>();
        int total = 0, expired = 0, expiringSoon = 0, longLived = 0;

        foreach (var c in credentials)
        {
            if (string.IsNullOrWhiteSpace(c.AppId) || string.IsNullOrWhiteSpace(c.DisplayName))
            {
                continue;
            }
            total++;

            if (!c.EndDateTime.HasValue)
            {
                continue;
            }

            var end = c.EndDateTime.Value;
            var label = string.IsNullOrWhiteSpace(c.CredentialDisplayName)
                ? c.CredentialType
                : $"{c.CredentialType} '{c.CredentialDisplayName}'";

            if (end < now)
            {
                expired++;
                var daysAgo = (int)Math.Floor((now - end).TotalDays);
                findings.Add(new AuditFinding(
                    "App credential expired",
                    c.DisplayName,
                    $"{label} expiró hace {daysAgo}d (endDate={end:yyyy-MM-dd}). AppId={c.AppId}",
                    "ERROR"));
                continue;
            }

            var daysToExpiry = (int)Math.Ceiling((end - now).TotalDays);
            if (daysToExpiry <= ExpiringSoonDays)
            {
                expiringSoon++;
                findings.Add(new AuditFinding(
                    "App credential expiring",
                    c.DisplayName,
                    $"{label} expira en {daysToExpiry}d (endDate={end:yyyy-MM-dd}). AppId={c.AppId}",
                    "WARN"));
                continue;
            }

            var lifespanDays = c.EndDateTime.HasValue
                ? (int)Math.Floor((end - now).TotalDays)
                : 0;
            if (lifespanDays > LongLivedDays)
            {
                longLived++;
                findings.Add(new AuditFinding(
                    "App credential long-lived",
                    c.DisplayName,
                    $"{label} caduca en {lifespanDays}d (>{LongLivedDays}d) — rotación poco frecuente. AppId={c.AppId}",
                    "INFO"));
            }
        }

        var summary = new AppCredentialsSummary(
            Total: total,
            Expired: expired,
            ExpiringSoon: expiringSoon,
            LongLived: longLived);

        return (summary, findings);
    }
}
