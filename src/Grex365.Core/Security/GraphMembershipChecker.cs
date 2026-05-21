using Grex365.Core.Abstractions;
using Microsoft.Graph.Me.CheckMemberGroups;

namespace Grex365.Core.Security;

public sealed class GraphMembershipChecker : IMembershipChecker
{
    private readonly IGraphConnection _connection;

    public GraphMembershipChecker(IGraphConnection connection)
    {
        _connection = connection;
    }

    public async Task<bool> IsMemberOfAsync(string groupId, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(groupId);

        var client = _connection.Client;
        if (client is null)
        {
            // Cert-based / app-only auth tiene no `me`; no podemos validar membership de un usuario.
            // Devolver false fuerza la denegación a la capa superior (consistente con "no auth user → no grant").
            return false;
        }

        var body = new CheckMemberGroupsPostRequestBody
        {
            GroupIds = new List<string> { groupId },
        };
        var response = await client.Me.CheckMemberGroups
            .PostAsCheckMemberGroupsPostResponseAsync(body, cancellationToken: cancellationToken)
            .ConfigureAwait(false);

        return response?.Value?.Contains(groupId, StringComparer.OrdinalIgnoreCase) == true;
    }
}
