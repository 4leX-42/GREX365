using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace Grex365.Core.Audit;

public sealed class GraphAuditService : IAuditService
{
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

        var analyzer = new IdentityAuditAnalyzer(DateTimeOffset.UtcNow);

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
                analyzer.Visit(ToSnapshot(user));
                return true;
            });

        await iterator.IterateAsync(cancellationToken).ConfigureAwait(false);

        var summary = analyzer.BuildSummary();
        progress?.Report(LogEntry.Ok("Audit", $"Procesados {summary.UsersTotal} usuarios; {analyzer.Findings.Count} hallazgos"));
        return (summary, analyzer.Findings);
    }

    public async Task<IReadOnlyList<AuditFinding>> RunGroupsAuditAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var client = _connection.Client
            ?? throw new InvalidOperationException("Graph no está conectado.");

        progress?.Report(LogEntry.Info("Audit", "Cargando grupos..."));
        var findings = new List<AuditFinding>();
        var processed = 0;

        var response = await client.Groups.GetAsync(req =>
        {
            req.QueryParameters.Select = new[]
            {
                "id", "displayName", "mail", "groupTypes", "mailEnabled", "securityEnabled"
            };
            req.QueryParameters.Top = 999;
            req.Headers.Add("ConsistencyLevel", "eventual");
        }, cancellationToken).ConfigureAwait(false);

        var iterator = PageIterator<Group, GroupCollectionResponse>.CreatePageIterator(
            client,
            response!,
            async group =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                await AnalyzeGroup(client, group, findings, cancellationToken).ConfigureAwait(false);
                processed++;
                return true;
            });

        await iterator.IterateAsync(cancellationToken).ConfigureAwait(false);
        progress?.Report(LogEntry.Ok("Audit", $"Procesados {processed} grupos; {findings.Count} hallazgos"));
        return findings;
    }

    private static async Task AnalyzeGroup(GraphServiceClient client, Group group, List<AuditFinding> findings, CancellationToken ct)
    {
        var id = group.Id ?? string.Empty;
        var name = group.DisplayName ?? "(sin nombre)";
        var types = group.GroupTypes ?? new List<string>();
        var isUnified = types.Contains("Unified");
        var isDl = group.MailEnabled == true && !isUnified && group.SecurityEnabled != true;

        try
        {
            var owners = await client.Groups[id].Owners.Count.GetAsync(req =>
            {
                req.Headers.Add("ConsistencyLevel", "eventual");
            }, ct).ConfigureAwait(false);
            if ((owners ?? 0) == 0)
            {
                findings.Add(new AuditFinding("Group without owner", name, $"id={id}", "WARN"));
            }
        }
        catch
        {
            // tolerate
        }

        if (isUnified || isDl)
        {
            try
            {
                var members = await client.Groups[id].Members.Count.GetAsync(req =>
                {
                    req.Headers.Add("ConsistencyLevel", "eventual");
                }, ct).ConfigureAwait(false);
                if ((members ?? 0) == 0)
                {
                    var category = isUnified ? "Empty M365 group" : "Empty DL";
                    findings.Add(new AuditFinding(category, name, $"id={id}", "INFO"));
                }
            }
            catch
            {
                // tolerate
            }
        }
    }

    private static UserSnapshot ToSnapshot(User u) => new(
        Id: u.Id,
        UserPrincipalName: u.UserPrincipalName,
        AccountEnabled: u.AccountEnabled ?? false,
        IsGuest: string.Equals(u.UserType, "Guest", StringComparison.OrdinalIgnoreCase),
        AssignedLicenseCount: u.AssignedLicenses?.Count ?? 0,
        LastSignIn: u.SignInActivity?.LastSignInDateTime);
}
