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

    public async Task<(PrivilegedRoleSummary Summary, IReadOnlyList<AuditFinding> Findings)> RunPrivilegedRolesAuditAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var client = _connection.Client
            ?? throw new InvalidOperationException("Graph no está conectado.");

        progress?.Report(LogEntry.Info("Audit", "Descargando directoryRoles activados..."));

        Microsoft.Graph.Models.DirectoryRoleCollectionResponse? rolesResp;
        try
        {
            rolesResp = await client.DirectoryRoles
                .GetAsync(cancellationToken: cancellationToken)
                .ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            progress?.Report(LogEntry.Error(
                "Audit",
                $"DirectoryRoles falló: {ex.Message}. ¿Directory.Read.All concedido?",
                ex));
            throw;
        }

        var roles = new List<DirectoryRole>();
        if (rolesResp is not null)
        {
            var rIter = PageIterator<DirectoryRole, Microsoft.Graph.Models.DirectoryRoleCollectionResponse>
                .CreatePageIterator(client, rolesResp, r =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    roles.Add(r);
                    return true;
                });
            await rIter.IterateAsync(cancellationToken).ConfigureAwait(false);
        }

        progress?.Report(LogEntry.Info("Audit", $"{roles.Count} roles activados — enumerando miembros..."));

        var assignments = new System.Collections.Concurrent.ConcurrentBag<PrivilegedRoleAssignment>();
        using var sem = new System.Threading.SemaphoreSlim(8);
        var tasks = roles.Select(async role =>
        {
            await sem.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                var roleId = role.Id ?? string.Empty;
                var roleName = role.DisplayName ?? "(sin nombre)";
                var template = role.RoleTemplateId;

                var memResp = await client.DirectoryRoles[roleId].Members
                    .GetAsync(req =>
                    {
                        req.QueryParameters.Select = new[]
                        {
                            "id", "displayName", "userPrincipalName", "userType", "accountEnabled"
                        };
                        req.QueryParameters.Top = 999;
                    }, cancellationToken)
                    .ConfigureAwait(false);
                if (memResp is null)
                {
                    return;
                }

                var mIter = PageIterator<DirectoryObject, DirectoryObjectCollectionResponse>
                    .CreatePageIterator(client, memResp, m =>
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        assignments.Add(ToAssignment(roleName, template, m));
                        return true;
                    });
                await mIter.IterateAsync(cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                progress?.Report(LogEntry.Warn("Audit",
                    $"No se pudo enumerar miembros de '{role.DisplayName}': {ex.Message}"));
            }
            finally
            {
                sem.Release();
            }
        });
        await Task.WhenAll(tasks).ConfigureAwait(false);

        var (summary, findings) = PrivilegedRoleAuditAnalyzer.Analyze(assignments);
        progress?.Report(LogEntry.Ok(
            "Audit",
            $"Privileged roles: {summary.GlobalAdmins} GA · {summary.UniqueAdmins} admins únicos · " +
            $"{summary.GuestsWithAdminRole} guests · {summary.DisabledWithAdminRole} disabled · " +
            $"{summary.ServicePrincipalsWithAdminRole} SP · {findings.Count} hallazgos"));
        return (summary, findings);
    }

    private static PrivilegedRoleAssignment ToAssignment(string roleName, string? template, DirectoryObject member)
    {
        var id = member.Id ?? string.Empty;
        return member switch
        {
            User u => new PrivilegedRoleAssignment(
                RoleName: roleName,
                RoleTemplateId: template,
                MemberId: id,
                MemberDisplayName: u.DisplayName,
                MemberUpn: u.UserPrincipalName,
                MemberType: "User",
                MemberUserType: u.UserType,
                MemberAccountEnabled: u.AccountEnabled ?? false),
            ServicePrincipal sp => new PrivilegedRoleAssignment(
                RoleName: roleName,
                RoleTemplateId: template,
                MemberId: id,
                MemberDisplayName: sp.DisplayName,
                MemberUpn: null,
                MemberType: "ServicePrincipal",
                MemberUserType: null,
                MemberAccountEnabled: sp.AccountEnabled ?? true),
            Group g => new PrivilegedRoleAssignment(
                RoleName: roleName,
                RoleTemplateId: template,
                MemberId: id,
                MemberDisplayName: g.DisplayName,
                MemberUpn: g.Mail,
                MemberType: "Group",
                MemberUserType: null,
                MemberAccountEnabled: true),
            _ => new PrivilegedRoleAssignment(
                RoleName: roleName,
                RoleTemplateId: template,
                MemberId: id,
                MemberDisplayName: member.OdataType,
                MemberUpn: null,
                MemberType: member.OdataType,
                MemberUserType: null,
                MemberAccountEnabled: true),
        };
    }

    public async Task<(AppCredentialsSummary Summary, IReadOnlyList<AuditFinding> Findings)> RunAppCredentialsAuditAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var client = _connection.Client
            ?? throw new InvalidOperationException("Graph no está conectado.");

        progress?.Report(LogEntry.Info("Audit", "Descargando /applications..."));

        Microsoft.Graph.Models.ApplicationCollectionResponse? response;
        try
        {
            response = await client.Applications.GetAsync(req =>
            {
                req.QueryParameters.Select = new[]
                {
                    "id", "appId", "displayName", "passwordCredentials", "keyCredentials"
                };
                req.QueryParameters.Top = 999;
            }, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            progress?.Report(LogEntry.Error(
                "Audit",
                $"Applications.Get falló: {ex.Message}. ¿Application.Read.All o Directory.Read.All concedido?",
                ex));
            throw;
        }

        var snapshots = new List<AppCredentialSnapshot>();
        if (response is not null)
        {
            var iter = PageIterator<Application, Microsoft.Graph.Models.ApplicationCollectionResponse>
                .CreatePageIterator(client, response, app =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var appId = app.AppId ?? string.Empty;
                    var name = app.DisplayName ?? "(sin nombre)";
                    if (app.PasswordCredentials is { } pwds)
                    {
                        foreach (var pwd in pwds)
                        {
                            snapshots.Add(new AppCredentialSnapshot(
                                AppId: appId,
                                DisplayName: name,
                                CredentialType: "Password",
                                KeyId: pwd.KeyId?.ToString(),
                                CredentialDisplayName: pwd.DisplayName,
                                EndDateTime: pwd.EndDateTime));
                        }
                    }
                    if (app.KeyCredentials is { } keys)
                    {
                        foreach (var key in keys)
                        {
                            snapshots.Add(new AppCredentialSnapshot(
                                AppId: appId,
                                DisplayName: name,
                                CredentialType: "Key",
                                KeyId: key.KeyId?.ToString(),
                                CredentialDisplayName: key.DisplayName,
                                EndDateTime: key.EndDateTime));
                        }
                    }
                    return true;
                });
            await iter.IterateAsync(cancellationToken).ConfigureAwait(false);
        }

        var (summary, findings) = AppCredentialAuditAnalyzer.Analyze(snapshots, DateTimeOffset.UtcNow);
        progress?.Report(LogEntry.Ok(
            "Audit",
            $"App creds: {summary.Total} totales · {summary.Expired} expired · " +
            $"{summary.ExpiringSoon} expiring · {summary.LongLived} long-lived · {findings.Count} hallazgos"));
        return (summary, findings);
    }

    public async Task<(TenantDefaultsSummary Summary, IReadOnlyList<AuditFinding> Findings)> RunTenantDefaultsAuditAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var client = _connection.Client
            ?? throw new InvalidOperationException("Graph no está conectado.");

        progress?.Report(LogEntry.Info("Audit", "Descargando /policies/authorizationPolicy..."));

        Microsoft.Graph.Models.AuthorizationPolicy? authPolicy;
        try
        {
            authPolicy = await client.Policies.AuthorizationPolicy
                .GetAsync(cancellationToken: cancellationToken)
                .ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            progress?.Report(LogEntry.Error(
                "Audit",
                $"AuthorizationPolicy falló: {ex.Message}. ¿Policy.Read.All concedido?",
                ex));
            throw;
        }

        if (authPolicy is null)
        {
            progress?.Report(LogEntry.Warn("Audit", "authorizationPolicy devolvió null."));
            return (new TenantDefaultsSummary(false, 0), Array.Empty<AuditFinding>());
        }

        bool securityDefaults = false;
        try
        {
            var sd = await client.Policies.IdentitySecurityDefaultsEnforcementPolicy
                .GetAsync(cancellationToken: cancellationToken)
                .ConfigureAwait(false);
            securityDefaults = sd?.IsEnabled == true;
        }
        catch (Exception ex)
        {
            progress?.Report(LogEntry.Warn("Audit",
                $"IdentitySecurityDefaultsEnforcementPolicy no disponible: {ex.Message}"));
        }

        var snapshot = new AuthorizationPolicySnapshot(
            AllowedToSignUpEmailBasedSubscriptions: authPolicy.AllowedToSignUpEmailBasedSubscriptions ?? false,
            AllowedToUseSspr: authPolicy.AllowedToUseSSPR ?? false,
            AllowEmailVerifiedUsersToJoinOrganization: authPolicy.AllowEmailVerifiedUsersToJoinOrganization ?? false,
            AllowInvitesFrom: authPolicy.AllowInvitesFrom?.ToString(),
            DefaultUserCanCreateApps: authPolicy.DefaultUserRolePermissions?.AllowedToCreateApps ?? false,
            DefaultUserCanCreateSecurityGroups: authPolicy.DefaultUserRolePermissions?.AllowedToCreateSecurityGroups ?? false,
            DefaultUserCanCreateTenants: authPolicy.DefaultUserRolePermissions?.AllowedToCreateTenants ?? false,
            DefaultUserCanReadOtherUsers: authPolicy.DefaultUserRolePermissions?.AllowedToReadOtherUsers ?? false);

        var (summary, findings) = TenantDefaultsAnalyzer.Analyze(snapshot, securityDefaults);
        progress?.Report(LogEntry.Ok(
            "Audit",
            $"Tenant defaults: SecurityDefaults={securityDefaults} · {findings.Count} hallazgos"));
        return (summary, findings);
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
