using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace Grex365.Core.Audit;

public sealed class GraphAuditService : IAuditService
{
    private static readonly TimeSpan MemberStaleAfter = TimeSpan.FromDays(180);
    private static readonly TimeSpan GuestStaleAfter = TimeSpan.FromDays(90);

    private readonly IGraphConnection _connection;

    public GraphAuditService(IGraphConnection connection)
    {
        _connection = connection;
    }

    public async Task<(AuditSummary Summary, IReadOnlyList<AuditFinding> Findings)> RunIdentityAuditAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var client = _connection.Client
            ?? throw new InvalidOperationException("Graph no está conectado.");

        progress?.Report(LogEntry.Info("Audit", "Cargando usuarios..."));

        var findings = new List<AuditFinding>();
        var totals = new Totals();
        var now = DateTimeOffset.UtcNow;

        var response = await client.Users.GetAsync(req =>
        {
            req.QueryParameters.Select = new[]
            {
                "id", "userPrincipalName", "displayName", "accountEnabled",
                "userType", "assignedLicenses", "signInActivity", "mail"
            };
            req.QueryParameters.Top = 999;
            req.Headers.Add("ConsistencyLevel", "eventual");
        }, cancellationToken).ConfigureAwait(false);

        var iterator = PageIterator<User, UserCollectionResponse>.CreatePageIterator(
            client,
            response!,
            user =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                ProcessUser(user, now, totals, findings);
                return true;
            });

        await iterator.IterateAsync(cancellationToken).ConfigureAwait(false);

        progress?.Report(LogEntry.Ok("Audit", $"Procesados {totals.UsersTotal} usuarios; {findings.Count} hallazgos"));

        var summary = new AuditSummary(
            UsersTotal: totals.UsersTotal,
            UsersEnabled: totals.UsersEnabled,
            UsersDisabled: totals.UsersDisabled,
            Guests: totals.Guests,
            StaleMembers: totals.StaleMembers,
            StaleGuests: totals.StaleGuests,
            DisabledWithLicense: totals.DisabledWithLicense);

        return (summary, findings);
    }

    private static void ProcessUser(User u, DateTimeOffset now, Totals totals, List<AuditFinding> findings)
    {
        totals.UsersTotal++;
        var upn = u.UserPrincipalName ?? u.Id ?? "(desconocido)";
        var isGuest = string.Equals(u.UserType, "Guest", StringComparison.OrdinalIgnoreCase);
        var enabled = u.AccountEnabled ?? false;
        var licenseCount = u.AssignedLicenses?.Count ?? 0;
        var lastSignIn = u.SignInActivity?.LastSignInDateTime;

        if (isGuest) totals.Guests++;
        if (enabled) totals.UsersEnabled++; else totals.UsersDisabled++;

        if (!enabled && licenseCount > 0)
        {
            totals.DisabledWithLicense++;
            findings.Add(new AuditFinding(
                "Disabled+License", upn,
                $"Deshabilitado con {licenseCount} licencias asignadas", "WARN"));
        }

        if (enabled)
        {
            var cutoff = isGuest ? now - GuestStaleAfter : now - MemberStaleAfter;
            if (lastSignIn is null || lastSignIn < cutoff)
            {
                var lastTxt = lastSignIn?.ToString("yyyy-MM-dd") ?? "nunca";
                var threshold = isGuest ? "90d" : "180d";
                if (isGuest)
                {
                    totals.StaleGuests++;
                    findings.Add(new AuditFinding(
                        "Stale guest", upn,
                        $"último login: {lastTxt} (>{threshold})", "WARN"));
                }
                else
                {
                    totals.StaleMembers++;
                    findings.Add(new AuditFinding(
                        "Stale member", upn,
                        $"último login: {lastTxt} (>{threshold})", "WARN"));
                }
            }
        }
    }

    private sealed class Totals
    {
        public int UsersTotal;
        public int UsersEnabled;
        public int UsersDisabled;
        public int Guests;
        public int StaleMembers;
        public int StaleGuests;
        public int DisabledWithLicense;
    }
}
