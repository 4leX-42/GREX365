using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Graph;

namespace Grex365.Core.Health;

public sealed class GraphTenantHealthService : ITenantHealthService
{
    private readonly IGraphConnection _connection;

    public GraphTenantHealthService(IGraphConnection connection)
    {
        _connection = connection;
    }

    public async Task<TenantHealth> GetAsync(IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default)
    {
        var client = _connection.Client
            ?? throw new InvalidOperationException("Graph no está conectado.");

        progress?.Report(LogEntry.Info("TenantHealth", "Cargando organización..."));
        var orgResponse = await client.Organization.GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
        var org = orgResponse?.Value?.FirstOrDefault();

        var tenantId = org?.Id ?? string.Empty;
        var name = org?.DisplayName ?? "(desconocido)";
        var verifiedDomain = org?.VerifiedDomains?.FirstOrDefault(d => d.IsDefault == true)?.Name;

        progress?.Report(LogEntry.Info("TenantHealth", "Contando usuarios y grupos..."));
        var usersCount = await CountAsync(client, "/users", cancellationToken).ConfigureAwait(false);
        var groupsCount = await CountAsync(client, "/groups", cancellationToken).ConfigureAwait(false);

        progress?.Report(LogEntry.Info("TenantHealth", "Cargando licencias..."));
        var skuResponse = await client.SubscribedSkus.GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
        var licenses = new List<LicenseSummary>();
        if (skuResponse?.Value is not null)
        {
            foreach (var sku in skuResponse.Value)
            {
                licenses.Add(new LicenseSummary(
                    SkuPartNumber: sku.SkuPartNumber ?? "?",
                    SkuId: sku.SkuId?.ToString() ?? string.Empty,
                    Consumed: sku.ConsumedUnits ?? 0,
                    Enabled: sku.PrepaidUnits?.Enabled ?? 0,
                    Warning: sku.PrepaidUnits?.Warning ?? 0,
                    Suspended: sku.PrepaidUnits?.Suspended ?? 0));
            }
        }

        return new TenantHealth(tenantId, name, verifiedDomain, usersCount, groupsCount, licenses);
    }

    private static async Task<int> CountAsync(GraphServiceClient client, string segment, CancellationToken ct)
    {
        try
        {
            if (segment == "/users")
            {
                var count = await client.Users.Count.GetAsync(req =>
                {
                    req.Headers.Add("ConsistencyLevel", "eventual");
                }, ct).ConfigureAwait(false);
                return count ?? 0;
            }
            if (segment == "/groups")
            {
                var count = await client.Groups.Count.GetAsync(req =>
                {
                    req.Headers.Add("ConsistencyLevel", "eventual");
                }, ct).ConfigureAwait(false);
                return count ?? 0;
            }
        }
        catch
        {
            // tolerate if /$count requires extra perms; return 0
        }
        return 0;
    }
}
