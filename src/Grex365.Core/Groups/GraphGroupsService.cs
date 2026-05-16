using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace Grex365.Core.Groups;

public sealed class GraphGroupsService : IGroupsService
{
    private readonly IGraphConnection _connection;

    public GraphGroupsService(IGraphConnection connection)
    {
        _connection = connection;
    }

    private GraphServiceClient RequireClient()
    {
        return _connection.Client
            ?? throw new InvalidOperationException("Graph no está conectado. Conecta antes de operar con grupos.");
    }

    public async Task<IReadOnlyList<GroupSummary>> SearchAsync(string query, CancellationToken cancellationToken = default)
    {
        var client = RequireClient();
        var trimmed = (query ?? string.Empty).Trim();

        var result = await client.Groups.GetAsync(req =>
        {
            req.QueryParameters.Select = ["id", "displayName", "mail", "groupTypes", "mailEnabled", "securityEnabled"];
            req.QueryParameters.Top = 50;
            if (!string.IsNullOrEmpty(trimmed))
            {
                var safe = trimmed.Replace("'", "''");
                req.QueryParameters.Filter =
                    $"startswith(displayName,'{safe}') or startswith(mail,'{safe}')";
            }
            req.Headers.Add("ConsistencyLevel", "eventual");
        }, cancellationToken).ConfigureAwait(false);

        var list = new List<GroupSummary>();
        if (result?.Value is null)
        {
            return list;
        }

        foreach (var g in result.Value)
        {
            list.Add(new GroupSummary(
                Id: g.Id ?? string.Empty,
                DisplayName: g.DisplayName ?? "(sin nombre)",
                Mail: g.Mail,
                GroupKind: ClassifyGroup(g)));
        }

        return list;
    }

    public async Task<IReadOnlyList<GroupMember>> GetMembersAsync(string groupId, CancellationToken cancellationToken = default)
    {
        var client = RequireClient();

        var result = await client.Groups[groupId].Members.GetAsync(req =>
        {
            req.QueryParameters.Top = 200;
            req.QueryParameters.Select = ["id", "displayName", "mail", "userPrincipalName"];
        }, cancellationToken).ConfigureAwait(false);

        var list = new List<GroupMember>();
        if (result?.Value is null)
        {
            return list;
        }

        foreach (var member in result.Value)
        {
            if (member is User u)
            {
                list.Add(new GroupMember(u.Id ?? string.Empty, u.DisplayName, u.Mail, u.UserPrincipalName));
            }
            else
            {
                list.Add(new GroupMember(member.Id ?? string.Empty, member.OdataType, null, null));
            }
        }
        return list;
    }

    public async Task<IReadOnlyList<AddMemberResult>> AddMembersAsync(
        string groupId,
        IReadOnlyCollection<string> userIdentifiers,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var client = RequireClient();
        var results = new List<AddMemberResult>(userIdentifiers.Count);

        foreach (var raw in userIdentifiers)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var input = (raw ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(input))
            {
                results.Add(new AddMemberResult(input, "VACIO", "Entrada vacía"));
                continue;
            }

            try
            {
                var userId = await ResolveUserIdAsync(client, input, cancellationToken).ConfigureAwait(false);
                if (userId is null)
                {
                    results.Add(new AddMemberResult(input, "NO_RESUELTO", "Usuario no encontrado en Graph"));
                    progress?.Report(LogEntry.Warn("Groups", $"No resuelto: {input}"));
                    continue;
                }

                var refBody = new Microsoft.Graph.Models.ReferenceCreate
                {
                    OdataId = $"https://graph.microsoft.com/v1.0/directoryObjects/{userId}"
                };
                await client.Groups[groupId].Members.Ref.PostAsync(refBody, cancellationToken: cancellationToken)
                    .ConfigureAwait(false);

                results.Add(new AddMemberResult(input, "AGREGADO", $"id={userId}"));
                progress?.Report(LogEntry.Ok("Groups", $"Agregado: {input}"));
            }
            catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
            {
                var code = ex.Error?.Code ?? "GraphError";
                if (string.Equals(code, "Request_BadRequest", StringComparison.OrdinalIgnoreCase)
                    && (ex.Error?.Message?.Contains("already exist", StringComparison.OrdinalIgnoreCase) ?? false))
                {
                    results.Add(new AddMemberResult(input, "YA_EXISTE", "Ya pertenece al grupo"));
                    progress?.Report(LogEntry.Info("Groups", $"Ya existe: {input}"));
                }
                else
                {
                    results.Add(new AddMemberResult(input, "ERROR", $"{code}: {ex.Error?.Message}"));
                    progress?.Report(LogEntry.Error("Groups", $"{input}: {ex.Error?.Message}", ex));
                }
            }
            catch (Exception ex)
            {
                results.Add(new AddMemberResult(input, "ERROR", ex.Message));
                progress?.Report(LogEntry.Error("Groups", $"{input}: {ex.Message}", ex));
            }
        }

        return results;
    }

    public async Task RemoveMemberAsync(
        string groupId,
        string memberId,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var client = RequireClient();
        try
        {
            await client.Groups[groupId].Members[memberId].Ref
                .DeleteAsync(cancellationToken: cancellationToken)
                .ConfigureAwait(false);
            progress?.Report(LogEntry.Ok("Groups", $"Eliminado del grupo: {memberId}"));
        }
        catch (Exception ex)
        {
            progress?.Report(LogEntry.Error("Groups", $"Error al eliminar {memberId}: {ex.Message}", ex));
            throw;
        }
    }

    public async Task<IReadOnlyList<BulkGroupResult>> CreateM365GroupsFromRowsAsync(
        IReadOnlyList<BulkGroupRow> rows,
        string domain,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var client = RequireClient();
        var results = new List<BulkGroupResult>();
        var cleanDomain = (domain ?? string.Empty).TrimStart('@').Trim();
        if (string.IsNullOrEmpty(cleanDomain))
        {
            throw new ArgumentException("Dominio requerido.", nameof(domain));
        }

        var groups = rows.GroupBy(r => r.GroupName, StringComparer.OrdinalIgnoreCase);
        foreach (var grp in groups)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var groupName = grp.Key.Trim();
            var groupEmail = $"{groupName}@{cleanDomain}";

            string? groupId;
            try
            {
                groupId = await EnsureM365GroupAsync(client, groupName, groupEmail, results, progress, cancellationToken)
                    .ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                results.Add(new BulkGroupResult(groupName, groupEmail, "Error", null, "Creación fallida: " + ex.Message));
                progress?.Report(LogEntry.Error("BulkGroups", $"{groupEmail}: {ex.Message}", ex));
                continue;
            }
            if (groupId is null)
            {
                continue;
            }

            var existing = await LoadM365MemberKeysAsync(client, groupId, cancellationToken).ConfigureAwait(false);

            foreach (var row in grp)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var email = row.Email.Trim();
                if (!BulkGroupRowPreprocessor.IsEmail(email))
                {
                    results.Add(new BulkGroupResult(groupName, groupEmail, "Error", email, "Email inválido"));
                    progress?.Report(LogEntry.Warn("BulkGroups", $"Email inválido: {email}"));
                    continue;
                }
                if (existing.Contains(email.ToLowerInvariant()))
                {
                    results.Add(new BulkGroupResult(groupName, groupEmail, "MemberSkipped", email, "Ya pertenece"));
                    continue;
                }

                try
                {
                    var userId = await ResolveUserIdAsync(client, email, cancellationToken).ConfigureAwait(false);
                    if (userId is null)
                    {
                        results.Add(new BulkGroupResult(groupName, groupEmail, "Error", email, "Usuario no encontrado"));
                        progress?.Report(LogEntry.Warn("BulkGroups", $"No resuelto: {email}"));
                        continue;
                    }

                    var refBody = new ReferenceCreate
                    {
                        OdataId = $"https://graph.microsoft.com/v1.0/directoryObjects/{userId}"
                    };
                    await client.Groups[groupId].Members.Ref.PostAsync(refBody, cancellationToken: cancellationToken)
                        .ConfigureAwait(false);
                    existing.Add(email.ToLowerInvariant());
                    results.Add(new BulkGroupResult(groupName, groupEmail, "MemberAdded", email, $"id={userId}"));
                    progress?.Report(LogEntry.Ok("BulkGroups", $"{groupEmail} + {email}"));
                }
                catch (Microsoft.Graph.Models.ODataErrors.ODataError ex) when (
                    string.Equals(ex.Error?.Code, "Request_BadRequest", StringComparison.OrdinalIgnoreCase)
                    && (ex.Error?.Message?.Contains("already exist", StringComparison.OrdinalIgnoreCase) ?? false))
                {
                    results.Add(new BulkGroupResult(groupName, groupEmail, "MemberSkipped", email, "Ya pertenece"));
                }
                catch (Exception ex)
                {
                    results.Add(new BulkGroupResult(groupName, groupEmail, "Error", email, "Adición fallida: " + ex.Message));
                    progress?.Report(LogEntry.Error("BulkGroups", $"{email}: {ex.Message}", ex));
                }
            }
        }
        return results;
    }

    private static async Task<string?> EnsureM365GroupAsync(
        GraphServiceClient client,
        string groupName,
        string groupEmail,
        List<BulkGroupResult> results,
        IProgress<LogEntry>? progress,
        CancellationToken cancellationToken)
    {
        var safe = groupEmail.Replace("'", "''");
        var existing = await client.Groups.GetAsync(req =>
        {
            req.QueryParameters.Filter = $"mail eq '{safe}'";
            req.QueryParameters.Select = ["id"];
            req.QueryParameters.Top = 1;
            req.Headers.Add("ConsistencyLevel", "eventual");
        }, cancellationToken).ConfigureAwait(false);

        var found = existing?.Value?.FirstOrDefault();
        if (found is not null)
        {
            results.Add(new BulkGroupResult(groupName, groupEmail, "Skipped", null, "Ya existía"));
            progress?.Report(LogEntry.Info("BulkGroups", $"Skip existente: {groupEmail}"));
            return found.Id;
        }

        var alias = groupEmail.Split('@', 2)[0];
        var body = new Group
        {
            DisplayName = groupName,
            MailNickname = alias,
            MailEnabled = true,
            SecurityEnabled = false,
            GroupTypes = ["Unified"],
            Visibility = "Private"
        };
        var created = await client.Groups.PostAsync(body, cancellationToken: cancellationToken).ConfigureAwait(false);
        results.Add(new BulkGroupResult(groupName, groupEmail, "Created", null, $"id={created?.Id}"));
        progress?.Report(LogEntry.Ok("BulkGroups", $"Creado: {groupEmail}"));
        return created?.Id;
    }

    private static async Task<HashSet<string>> LoadM365MemberKeysAsync(
        GraphServiceClient client, string groupId, CancellationToken cancellationToken)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var resp = await client.Groups[groupId].Members.GetAsync(req =>
        {
            req.QueryParameters.Top = 200;
            req.QueryParameters.Select = ["id", "mail", "userPrincipalName"];
        }, cancellationToken).ConfigureAwait(false);
        if (resp?.Value is null)
        {
            return set;
        }
        foreach (var m in resp.Value)
        {
            if (m is User u)
            {
                if (!string.IsNullOrEmpty(u.Mail)) set.Add(u.Mail.ToLowerInvariant());
                if (!string.IsNullOrEmpty(u.UserPrincipalName)) set.Add(u.UserPrincipalName.ToLowerInvariant());
            }
        }
        return set;
    }

    private static async Task<string?> ResolveUserIdAsync(GraphServiceClient client, string input, CancellationToken cancellationToken)
    {
        if (Guid.TryParse(input, out _))
        {
            try
            {
                var user = await client.Users[input].GetAsync(req =>
                {
                    req.QueryParameters.Select = ["id"];
                }, cancellationToken).ConfigureAwait(false);
                if (user?.Id is not null)
                {
                    return user.Id;
                }
            }
            catch
            {
                // fall through to email-based lookup
            }
        }

        var safe = input.Replace("'", "''");
        var users = await client.Users.GetAsync(req =>
        {
            req.QueryParameters.Filter = $"mail eq '{safe}' or userPrincipalName eq '{safe}'";
            req.QueryParameters.Select = ["id"];
            req.QueryParameters.Top = 1;
            req.Headers.Add("ConsistencyLevel", "eventual");
        }, cancellationToken).ConfigureAwait(false);

        return users?.Value?.FirstOrDefault()?.Id;
    }

    private static string ClassifyGroup(Group g)
    {
        var types = g.GroupTypes ?? new List<string>();
        if (types.Contains("Unified"))
        {
            return "M365";
        }
        if (g.MailEnabled == true && g.SecurityEnabled == true)
        {
            return "MailSecurity";
        }
        if (g.MailEnabled == true)
        {
            return "DistributionList";
        }
        if (g.SecurityEnabled == true)
        {
            return "Security";
        }
        return "Other";
    }
}
