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
