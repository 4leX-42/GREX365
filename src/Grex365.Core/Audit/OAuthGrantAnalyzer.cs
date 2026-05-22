using Grex365.Core.Models;

namespace Grex365.Core.Audit;

public sealed record OAuthGrantSnapshot(
    string GrantId,
    string ClientId,
    string ClientDisplayName,
    string ResourceId,
    string ResourceDisplayName,
    string ConsentType,    // "AllPrincipals" | "Principal"
    string? PrincipalId,
    IReadOnlyList<string> Scopes);

public sealed record OAuthGrantsSummary(
    int TotalGrants,
    int TenantWideHighRisk,
    int UserConsentedHighRisk,
    int UniqueClients);

public static class OAuthGrantAnalyzer
{
    private static readonly HashSet<string> HighRiskScopes = new(StringComparer.OrdinalIgnoreCase)
    {
        "Mail.Read",
        "Mail.ReadWrite",
        "Mail.Read.Shared",
        "Mail.ReadWrite.Shared",
        "Mail.Send",
        "Mail.Send.Shared",
        "MailboxSettings.ReadWrite",
        "Files.Read.All",
        "Files.ReadWrite.All",
        "Sites.Read.All",
        "Sites.ReadWrite.All",
        "Sites.FullControl.All",
        "Sites.Manage.All",
        "Directory.Read.All",
        "Directory.ReadWrite.All",
        "Directory.AccessAsUser.All",
        "User.Read.All",
        "User.ReadWrite.All",
        "Group.Read.All",
        "Group.ReadWrite.All",
        "full_access_as_user",
        "Calendars.ReadWrite",
        "Calendars.ReadWrite.Shared",
        "Contacts.ReadWrite",
        "Notes.ReadWrite.All",
    };

    public static (OAuthGrantsSummary Summary, IReadOnlyList<AuditFinding> Findings) Analyze(
        IEnumerable<OAuthGrantSnapshot> grants)
    {
        ArgumentNullException.ThrowIfNull(grants);

        var findings = new List<AuditFinding>();
        var clientIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int total = 0, tenantWide = 0, userConsented = 0;

        foreach (var g in grants)
        {
            if (string.IsNullOrWhiteSpace(g.ClientId))
            {
                continue;
            }
            total++;
            clientIds.Add(g.ClientId);

            var highRiskHit = g.Scopes
                .Where(s => HighRiskScopes.Contains(s.Trim()))
                .ToList();
            if (highRiskHit.Count == 0)
            {
                continue;
            }

            var client = !string.IsNullOrWhiteSpace(g.ClientDisplayName)
                ? g.ClientDisplayName
                : g.ClientId;
            var scopeList = string.Join(", ", highRiskHit);
            var isTenantWide = string.Equals(g.ConsentType, "AllPrincipals", StringComparison.OrdinalIgnoreCase);

            if (isTenantWide)
            {
                tenantWide++;
                findings.Add(new AuditFinding(
                    "OAuth grant tenant-wide high-risk",
                    client,
                    $"Admin consent (AllPrincipals) sobre '{g.ResourceDisplayName}' con scopes peligrosos: {scopeList}. ClientId={g.ClientId}",
                    "ERROR"));
            }
            else
            {
                userConsented++;
                var who = g.PrincipalId ?? "(usuario)";
                findings.Add(new AuditFinding(
                    "OAuth grant user-consented high-risk",
                    client,
                    $"Consent de usuario {who} sobre '{g.ResourceDisplayName}' con scopes peligrosos: {scopeList}. Posible phishing OAuth. ClientId={g.ClientId}",
                    "WARN"));
            }
        }

        var summary = new OAuthGrantsSummary(
            TotalGrants: total,
            TenantWideHighRisk: tenantWide,
            UserConsentedHighRisk: userConsented,
            UniqueClients: clientIds.Count);

        return (summary, findings);
    }

    public static bool IsHighRiskScope(string scope) =>
        !string.IsNullOrWhiteSpace(scope) && HighRiskScopes.Contains(scope.Trim());
}
