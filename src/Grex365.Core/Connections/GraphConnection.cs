using System.Security.Cryptography.X509Certificates;
using Azure.Identity;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace Grex365.Core.Connections;

public sealed class GraphConnection : IGraphConnection
{
    private GraphServiceClient? _client;

    public bool IsConnected => _client is not null;

    public GraphServiceClient? Client => _client;

    public async Task ConnectByCertificateAsync(
        CertConfig config,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        progress?.Report(LogEntry.Info("Graph", $"Conectando a Microsoft Graph (cert) tenant={config.TenantId}"));

        var cert = LoadCertificateFromStore(config.CertThumbprint)
            ?? throw new InvalidOperationException(
                $"Certificado con thumbprint '{config.CertThumbprint}' no encontrado en Cert:\\CurrentUser\\My.");

        var credential = new ClientCertificateCredential(config.TenantId, config.AppId, cert);
        var client = new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);

        // Smoke test: read organization. Confirms cert auth works before we mark connected.
        Organization? org = null;
        try
        {
            var orgResponse = await client.Organization
                .GetAsync(cancellationToken: cancellationToken)
                .ConfigureAwait(false);
            org = orgResponse?.Value?.FirstOrDefault();
        }
        catch (Exception ex)
        {
            progress?.Report(LogEntry.Error("Graph", $"Smoke test falló: {ex.Message}", ex));
            throw;
        }

        _client = client;
        progress?.Report(LogEntry.Ok("Graph", $"Conectado. Organización: {org?.DisplayName ?? "?"}"));
    }

    public Task DisconnectAsync(CancellationToken cancellationToken = default)
    {
        _client = null;
        return Task.CompletedTask;
    }

    private static X509Certificate2? LoadCertificateFromStore(string thumbprint)
    {
        using var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        store.Open(OpenFlags.ReadOnly);
        var certs = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, validOnly: false);
        return certs.Count > 0 ? certs[0] : null;
    }
}
