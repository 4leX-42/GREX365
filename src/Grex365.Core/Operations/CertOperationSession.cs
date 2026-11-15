using Grex365.Core.Abstractions;

namespace Grex365.Core.Operations;

// Session for headless hosts (CLI, MCP server, scheduler): connects with the certificate the
// operator configured in the desktop app and enforces the same tenant lock.
public sealed class CertOperationSession : IOperationSession
{
    private readonly IGraphConnection _graph;
    private readonly ICertConfigStore _certStore;
    private readonly ITenantLock _tenantLock;
    private readonly SemaphoreSlim _gate = new(1, 1);

    public CertOperationSession(IGraphConnection graph, ICertConfigStore certStore, ITenantLock tenantLock)
    {
        _graph = graph;
        _certStore = certStore;
        _tenantLock = tenantLock;
    }

    public async Task EnsureAsync(OperationDescriptor descriptor, CancellationToken cancellationToken)
    {
        if (!descriptor.RequiresGraph && !descriptor.RequiresExchange)
        {
            return;
        }

        var config = await _certStore.LoadAsync(cancellationToken).ConfigureAwait(false)
            ?? throw new InvalidOperationException(
                "no hay conexión por certificado configurada. Abre GREX365 y completa el asistente de certificado.");

        // Exchange operations connect per call in an external PowerShell host with this same
        // configuration, so only Graph keeps a session here.
        if (!descriptor.RequiresGraph || _graph.IsConnected)
        {
            return;
        }

        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (_graph.IsConnected)
            {
                return;
            }

            await _graph.ConnectByCertificateAsync(config, null, cancellationToken).ConfigureAwait(false);
            try
            {
                await _tenantLock.EnforceAsync(_graph.TenantId ?? string.Empty, cancellationToken).ConfigureAwait(false);
            }
            catch
            {
                await _graph.DisconnectAsync(cancellationToken).ConfigureAwait(false);
                throw;
            }
        }
        finally
        {
            _gate.Release();
        }
    }
}
