namespace Grex365.Core.Models;

public sealed record ConnectionState(
    bool GraphConnected,
    bool ExchangeConnected,
    string? TenantId,
    string? TenantDomain,
    string? Account)
{
    public static ConnectionState Disconnected { get; } =
        new(false, false, null, null, null);

    public bool BothConnected => GraphConnected && ExchangeConnected;
}
