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

    public async Task<(MfaCoverageSummary Summary, IReadOnlyList<AuditFinding> Findings)> RunMfaCoverageAuditAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var client = _connection.Client
            ?? throw new InvalidOperationException("Graph no está conectado.");

        progress?.Report(LogEntry.Info("Audit", "Descargando userRegistrationDetails..."));

        Microsoft.Graph.Models.UserRegistrationDetailsCollectionResponse? response;
        try
        {
            response = await client.Reports.AuthenticationMethods.UserRegistrationDetails
                .GetAsync(req => req.QueryParameters.Top = 999, cancellationToken)
                .ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            progress?.Report(LogEntry.Error(
                "Audit",
                $"userRegistrationDetails falló: {ex.Message}. ¿Reports.Read.All concedido?",
                ex));
            throw;
        }

        var rows = new List<MfaRegistrationRow>();
        if (response is not null)
        {
            var iterator = Microsoft.Graph.PageIterator<
                    Microsoft.Graph.Models.UserRegistrationDetails,
                    Microsoft.Graph.Models.UserRegistrationDetailsCollectionResponse>
                .CreatePageIterator(
                    client,
                    response,
                    detail =>
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        rows.Add(new MfaRegistrationRow(
                            UserPrincipalName: detail.UserPrincipalName ?? string.Empty,
                            DisplayName: detail.UserDisplayName,
                            UserType: detail.UserType?.ToString(),
                            IsAdmin: detail.IsAdmin == true,
                            IsMfaRegistered: detail.IsMfaRegistered == true,
                            IsMfaCapable: detail.IsMfaCapable == true));
                        return true;
                    });
            await iterator.IterateAsync(cancellationToken).ConfigureAwait(false);
        }

        var (summary, findings) = MfaCoverageAnalyzer.Analyze(rows);
        progress?.Report(LogEntry.Ok(
            "Audit",
            $"MFA: {summary.AdminsTotal} admins ({summary.AdminsWithoutMfa} sin MFA), " +
            $"{summary.MembersTotal} miembros ({summary.MembersWithoutMfa} sin MFA), " +
            $"{summary.GuestsTotal} invitados ({summary.GuestsWithoutMfa} sin MFA)"));
        return (summary, findings);
    }

    public async Task<(CaPoliciesSummary Summary, IReadOnlyList<AuditFinding> Findings)> RunConditionalAccessAuditAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var client = _connection.Client
            ?? throw new InvalidOperationException("Graph no está conectado.");

        progress?.Report(LogEntry.Info("Audit", "Descargando Conditional Access policies..."));

        Microsoft.Graph.Models.ConditionalAccessPolicyCollectionResponse? response;
        try
        {
            response = await client.Identity.ConditionalAccess.Policies
                .GetAsync(cancellationToken: cancellationToken)
                .ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            progress?.Report(LogEntry.Error(
                "Audit",
                $"ConditionalAccess.Policies falló: {ex.Message}. ¿Policy.Read.All concedido?",
                ex));
            throw;
        }

        var snapshots = new List<CaPolicySnapshot>();
        if (response is not null)
        {
            var iterator = Microsoft.Graph.PageIterator<
                    Microsoft.Graph.Models.ConditionalAccessPolicy,
                    Microsoft.Graph.Models.ConditionalAccessPolicyCollectionResponse>
                .CreatePageIterator(
                    client,
                    response,
                    policy =>
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        snapshots.Add(ToSnapshot(policy));
                        return true;
                    });
            await iterator.IterateAsync(cancellationToken).ConfigureAwait(false);
        }

        var (summary, findings) = CaPolicyAnalyzer.Analyze(snapshots, DateTimeOffset.UtcNow);
        progress?.Report(LogEntry.Ok(
            "Audit",
            $"CA: {summary.Total} policies · {summary.Enabled} enabled · " +
            $"{summary.Disabled} disabled · {summary.ReportOnly} report-only · " +
            $"{findings.Count} hallazgos"));
        return (summary, findings);
    }

    private static CaPolicySnapshot ToSnapshot(Microsoft.Graph.Models.ConditionalAccessPolicy p)
    {
        var users = p.Conditions?.Users;
        var apps = p.Conditions?.Applications;
        var grant = p.GrantControls;

        return new CaPolicySnapshot(
            Id: p.Id ?? string.Empty,
            DisplayName: p.DisplayName ?? string.Empty,
            State: p.State?.ToString() ?? string.Empty,
            CreatedDateTime: p.CreatedDateTime,
            ModifiedDateTime: p.ModifiedDateTime,
            IncludeUsers: users?.IncludeUsers?.ToList() ?? new List<string>(),
            ExcludeUsers: users?.ExcludeUsers?.ToList() ?? new List<string>(),
            IncludeGroups: users?.IncludeGroups?.ToList() ?? new List<string>(),
            ExcludeGroups: users?.ExcludeGroups?.ToList() ?? new List<string>(),
            IncludeRoles: users?.IncludeRoles?.ToList() ?? new List<string>(),
            ExcludeRoles: users?.ExcludeRoles?.ToList() ?? new List<string>(),
            IncludeApplications: apps?.IncludeApplications?.ToList() ?? new List<string>(),
            BuiltInControls: grant?.BuiltInControls?.Select(c => c.ToString() ?? string.Empty).ToList()
                ?? new List<string>(),
            GrantOperator: grant?.Operator);
    }

    public async Task<IReadOnlyList<AuditFinding>> RunGroupActivityAuditAsync(
        int inactivityDays = 90,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (inactivityDays < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(inactivityDays), "Debe ser >= 1.");
        }

        var client = _connection.Client
            ?? throw new InvalidOperationException("Graph no está conectado.");

        // Map inactivityDays to the closest Graph report period bucket (D7/D30/D90/D180).
        var period = inactivityDays switch
        {
            <= 7 => "D7",
            <= 30 => "D30",
            <= 90 => "D90",
            _ => "D180",
        };

        progress?.Report(LogEntry.Info("Audit", $"Descargando reporte actividad grupos (period={period})..."));

        Stream? csvStream;
        try
        {
            csvStream = await client.Reports
                .GetOffice365GroupsActivityDetailWithPeriod(period)
                .GetAsync(cancellationToken: cancellationToken)
                .ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            progress?.Report(LogEntry.Error(
                "Audit",
                $"Reports.GetOffice365GroupsActivityDetail falló: {ex.Message}. ¿Permiso 'Reports.Read.All' concedido?",
                ex));
            throw;
        }

        if (csvStream is null)
        {
            progress?.Report(LogEntry.Warn("Audit", "El reporte devolvió cuerpo vacío."));
            return Array.Empty<AuditFinding>();
        }

        IReadOnlyList<GroupActivityRow> rows;
        await using (csvStream.ConfigureAwait(false))
        {
            rows = GroupActivityAnalyzer.ParseCsv(csvStream);
        }

        var today = DateOnly.FromDateTime(DateTime.UtcNow);
        var findings = GroupActivityAnalyzer.Analyze(rows, today, inactivityDays);
        progress?.Report(LogEntry.Ok(
            "Audit",
            $"Procesados {rows.Count} grupos en reporte; {findings.Count} inactivos (>{inactivityDays}d)"));
        return findings;
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
