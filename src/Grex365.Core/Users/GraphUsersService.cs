using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.AssignLicense;

namespace Grex365.Core.Users;

public sealed class GraphUsersService : IUsersService
{
    private readonly IGraphConnection _connection;

    public GraphUsersService(IGraphConnection connection)
    {
        _connection = connection;
    }

    private GraphServiceClient Client => _connection.Client
        ?? throw new InvalidOperationException("Graph no está conectado.");

    public async Task<IReadOnlyList<UserSummary>> SearchAsync(string query, CancellationToken cancellationToken = default)
    {
        var trimmed = (query ?? string.Empty).Trim();

        var response = await Client.Users.GetAsync(req =>
        {
            req.QueryParameters.Select = new[]
            {
                "id", "displayName", "userPrincipalName", "mail",
                "accountEnabled", "userType", "assignedLicenses", "signInActivity"
            };
            req.QueryParameters.Top = 50;
            if (!string.IsNullOrEmpty(trimmed))
            {
                var safe = trimmed.Replace("'", "''");
                req.QueryParameters.Filter =
                    $"startswith(displayName,'{safe}') or startswith(userPrincipalName,'{safe}') or startswith(mail,'{safe}')";
            }
            req.Headers.Add("ConsistencyLevel", "eventual");
        }, cancellationToken).ConfigureAwait(false);

        var list = new List<UserSummary>();
        if (response?.Value is null)
        {
            return list;
        }

        foreach (var u in response.Value)
        {
            list.Add(Map(u));
        }
        return list;
    }

    public async Task<UserSummary?> GetByIdAsync(string id, CancellationToken cancellationToken = default)
    {
        try
        {
            var u = await Client.Users[id].GetAsync(req =>
            {
                req.QueryParameters.Select = new[]
                {
                    "id", "displayName", "userPrincipalName", "mail",
                    "accountEnabled", "userType", "assignedLicenses", "signInActivity"
                };
            }, cancellationToken).ConfigureAwait(false);
            return u is null ? null : Map(u);
        }
        catch
        {
            return null;
        }
    }

    public async Task<IReadOnlyList<GroupSummary>> GetGroupMembershipsAsync(string userId, CancellationToken cancellationToken = default)
    {
        var response = await Client.Users[userId].MemberOf.GetAsync(req =>
        {
            req.QueryParameters.Top = 200;
        }, cancellationToken).ConfigureAwait(false);

        var list = new List<GroupSummary>();
        if (response?.Value is null)
        {
            return list;
        }

        foreach (var obj in response.Value)
        {
            if (obj is Group g)
            {
                list.Add(new GroupSummary(
                    Id: g.Id ?? string.Empty,
                    DisplayName: g.DisplayName ?? "(sin nombre)",
                    Mail: g.Mail,
                    GroupKind: ClassifyGroup(g)));
            }
        }
        return list;
    }

    public async Task SetAccountEnabledAsync(string userId, bool enabled, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default)
    {
        var body = new User { AccountEnabled = enabled };
        await Client.Users[userId].PatchAsync(body, cancellationToken: cancellationToken).ConfigureAwait(false);
        progress?.Report(LogEntry.Ok("Users", enabled ? $"Habilitado: {userId}" : $"Deshabilitado: {userId}"));
    }

    public async Task RemoveAllLicensesAsync(string userId, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default)
    {
        var user = await Client.Users[userId].GetAsync(req =>
        {
            req.QueryParameters.Select = new[] { "id", "assignedLicenses" };
        }, cancellationToken).ConfigureAwait(false);

        var skuIds = user?.AssignedLicenses?
            .Where(l => l.SkuId.HasValue)
            .Select(l => (Guid?)l.SkuId!.Value)
            .ToList() ?? new List<Guid?>();

        if (skuIds.Count == 0)
        {
            progress?.Report(LogEntry.Info("Users", $"Sin licencias que quitar para {userId}"));
            return;
        }

        var body = new AssignLicensePostRequestBody
        {
            AddLicenses = new List<AssignedLicense>(),
            RemoveLicenses = skuIds
        };
        await Client.Users[userId].AssignLicense.PostAsync(body, cancellationToken: cancellationToken).ConfigureAwait(false);
        progress?.Report(LogEntry.Ok("Users", $"Quitadas {skuIds.Count} licencias de {userId}"));
    }

    public async Task AssignLicenseAsync(string userId, Guid skuId, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default)
    {
        var body = new AssignLicensePostRequestBody
        {
            AddLicenses = new List<AssignedLicense>
            {
                new() { SkuId = skuId }
            },
            RemoveLicenses = new List<Guid?>()
        };
        await Client.Users[userId].AssignLicense.PostAsync(body, cancellationToken: cancellationToken).ConfigureAwait(false);
        progress?.Report(LogEntry.Ok("Users", $"Asignada licencia {skuId} a {userId}"));
    }

    public async Task<IReadOnlyList<SkuInfo>> ListSkusAsync(CancellationToken cancellationToken = default)
    {
        var response = await Client.SubscribedSkus.GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
        var list = new List<SkuInfo>();
        if (response?.Value is null)
        {
            return list;
        }
        foreach (var sku in response.Value)
        {
            if (!sku.SkuId.HasValue)
            {
                continue;
            }
            list.Add(new SkuInfo(
                SkuId: sku.SkuId.Value,
                SkuPartNumber: sku.SkuPartNumber ?? "?",
                Enabled: sku.PrepaidUnits?.Enabled ?? 0,
                Consumed: sku.ConsumedUnits ?? 0));
        }
        return list
            .OrderByDescending(s => s.Available)
            .ThenBy(s => s.SkuPartNumber)
            .ToList();
    }

    private static UserSummary Map(User u) => new(
        Id: u.Id ?? string.Empty,
        DisplayName: u.DisplayName,
        UserPrincipalName: u.UserPrincipalName,
        Mail: u.Mail,
        AccountEnabled: u.AccountEnabled ?? false,
        IsGuest: string.Equals(u.UserType, "Guest", StringComparison.OrdinalIgnoreCase),
        AssignedLicenseCount: u.AssignedLicenses?.Count ?? 0,
        LastSignIn: u.SignInActivity?.LastSignInDateTime);

    private static string ClassifyGroup(Group g)
    {
        var types = g.GroupTypes ?? new List<string>();
        if (types.Contains("Unified")) return "M365";
        if (g.MailEnabled == true && g.SecurityEnabled == true) return "MailSecurity";
        if (g.MailEnabled == true) return "DistributionList";
        if (g.SecurityEnabled == true) return "Security";
        return "Other";
    }
}
