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
        var findings = new System.Collections.Concurrent.ConcurrentBag<AuditFinding>();
        var groups = new List<Group>();

        var response = await client.Groups.GetAsync(req =>
        {
            req.QueryParameters.Select = new[]
            {
                "id", "displayName", "mail", "groupTypes", "mailEnabled", "securityEnabled", "visibility"
            };
            req.QueryParameters.Top = 999;
            req.Headers.Add("ConsistencyLevel", "eventual");
        }, cancellationToken).ConfigureAwait(false);

        var iterator = PageIterator<Group, GroupCollectionResponse>.CreatePageIterator(
            client,
            response!,
            group =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                groups.Add(group);
                return true;
            });
        await iterator.IterateAsync(cancellationToken).ConfigureAwait(false);

        progress?.Report(LogEntry.Info("Audit", $"Analizando {groups.Count} grupos en paralelo..."));

        using var sem = new System.Threading.SemaphoreSlim(8);
        var tasks = groups.Select(async g =>
        {
            await sem.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                await AnalyzeGroup(client, g, findings, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                sem.Release();
            }
        });
        await Task.WhenAll(tasks).ConfigureAwait(false);

        var sorted = findings.OrderBy(f => f.Category).ThenBy(f => f.Identity).ToList();
        progress?.Report(LogEntry.Ok("Audit", $"Procesados {groups.Count} grupos; {sorted.Count} hallazgos"));
        return sorted;
    }

    private static async Task AnalyzeGroup(GraphServiceClient client, Group group, System.Collections.Concurrent.ConcurrentBag<AuditFinding> findings, CancellationToken ct)
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

        if (isUnified && string.Equals(group.Visibility, "Private", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                var membersResp = await client.Groups[id].Members.GetAsync(req =>
                {
                    req.QueryParameters.Select = new[] { "id", "displayName", "userType", "userPrincipalName", "mail" };
                    req.QueryParameters.Top = 999;
                }, ct).ConfigureAwait(false);
                var guestIterator = PageIterator<DirectoryObject, DirectoryObjectCollectionResponse>.CreatePageIterator(
                    client,
                    membersResp!,
                    member =>
                    {
                        if (member is User u && string.Equals(u.UserType, "Guest", StringComparison.OrdinalIgnoreCase))
                        {
                            var who = u.UserPrincipalName ?? u.Mail ?? u.DisplayName ?? u.Id ?? "?";
                            findings.Add(new AuditFinding(
                                "Guest in private M365 group",
                                $"{name} ← {who}",
                                $"groupId={id}; userId={u.Id}",
                                "WARN"));
                        }
                        return true;
                    });
                await guestIterator.IterateAsync(ct).ConfigureAwait(false);
            }
            catch
            {
                // tolerate (Graph perms may not allow expanding members for some groups)
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
