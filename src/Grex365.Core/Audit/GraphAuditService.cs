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

    private static UserSnapshot ToSnapshot(User u) => new(
        Id: u.Id,
        UserPrincipalName: u.UserPrincipalName,
        AccountEnabled: u.AccountEnabled ?? false,
        IsGuest: string.Equals(u.UserType, "Guest", StringComparison.OrdinalIgnoreCase),
        AssignedLicenseCount: u.AssignedLicenses?.Count ?? 0,
        LastSignIn: u.SignInActivity?.LastSignInDateTime);
}
