using System.Security.Cryptography.X509Certificates;
using Azure.Identity;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace Grex365.Core.Connections;

public sealed class GraphConnection : IGraphConnection
{
    // Azure CLI's public-client app — multi-tenant, allows device-code auth without pre-registration.
    // Same trick used by Microsoft.Graph PowerShell module and Az CLI.
    private const string DeviceCodeClientId = "04b07795-8ddb-461a-bbee-02f9e1bf7b46";

    private GraphServiceClient? _client;
    private string? _tenantId;
    private string? _account;
    private DateTimeOffset _lastProbe = DateTimeOffset.MinValue;
    private bool _lastProbeResult;
    private static readonly TimeSpan ProbeCacheTtl = TimeSpan.FromSeconds(10);

    public bool IsConnected => _client is not null;

    public GraphServiceClient? Client => _client;

    public string? TenantId => _tenantId;

    public string? Account => _account;

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
        _tenantId = config.TenantId;
        _account = $"App ({config.AppId})";
        _lastProbe = DateTimeOffset.Now;
        _lastProbeResult = true;
        progress?.Report(LogEntry.Ok("Graph", $"Conectado. Organización: {org?.DisplayName ?? "?"}"));
    }

    public async Task ConnectByDeviceCodeAsync(
        string? tenantHint,
        Func<DeviceCodePrompt, CancellationToken, Task> codeCallback,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(codeCallback);

        var tenant = string.IsNullOrWhiteSpace(tenantHint) ? "organizations" : tenantHint;
        progress?.Report(LogEntry.Info("Graph", $"Iniciando device-code auth tenant={tenant}..."));

        var options = new DeviceCodeCredentialOptions
        {
            ClientId = DeviceCodeClientId,
            TenantId = tenant,
            DisableAutomaticAuthentication = false,
            DeviceCodeCallback = async (info, ct) =>
            {
                var prompt = new DeviceCodePrompt(
                    UserCode: info.UserCode,
                    VerificationUri: info.VerificationUri.ToString(),
                    Message: info.Message,
                    ExpiresOn: info.ExpiresOn);
                progress?.Report(LogEntry.Info("Graph", $"Codigo: {info.UserCode} — abre {info.VerificationUri}"));
                await codeCallback(prompt, ct).ConfigureAwait(false);
            }
        };

        var credential = new DeviceCodeCredential(options);
        var client = new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);

        Organization? org;
        User? me;
        try
        {
            var orgResponse = await client.Organization
                .GetAsync(cancellationToken: cancellationToken)
                .ConfigureAwait(false);
            org = orgResponse?.Value?.FirstOrDefault();

            me = await client.Me
                .GetAsync(cancellationToken: cancellationToken)
                .ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            progress?.Report(LogEntry.Error("Graph", $"Auth fallo: {ex.Message}", ex));
            throw;
        }

        _client = client;
        _tenantId = org?.Id;
        _account = me?.UserPrincipalName ?? me?.DisplayName ?? "(device-code)";
        _lastProbe = DateTimeOffset.Now;
        _lastProbeResult = true;
        progress?.Report(LogEntry.Ok("Graph", $"Conectado como {_account} a {org?.DisplayName ?? "?"}"));
    }

    public async Task<bool> CheckLiveAsync(CancellationToken cancellationToken = default)
    {
        if (_client is null)
        {
            return false;
        }

        var now = DateTimeOffset.Now;
        if (now - _lastProbe < ProbeCacheTtl)
        {
            return _lastProbeResult;
        }

        try
        {
            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(TimeSpan.FromSeconds(5));

            var orgResponse = await _client.Organization
                .GetAsync(cancellationToken: timeoutCts.Token)
                .ConfigureAwait(false);
            _lastProbeResult = orgResponse?.Value?.Count > 0;
        }
        catch
        {
            _lastProbeResult = false;
        }

        _lastProbe = now;
        return _lastProbeResult;
    }

    public Task DisconnectAsync(CancellationToken cancellationToken = default)
    {
        _client = null;
        _tenantId = null;
        _account = null;
        _lastProbeResult = false;
        _lastProbe = DateTimeOffset.MinValue;
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
