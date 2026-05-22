using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Graph.Models;

namespace Grex365.Core.Connections;

public sealed class GraphAppRegistrationService : IAppRegistrationService
{
    private readonly IGraphConnection _connection;

    public GraphAppRegistrationService(IGraphConnection connection)
    {
        _connection = connection;
    }

    public async Task<AppRegistrationResult> CreateAndConfigureAsync(
        string displayName,
        byte[] certificateCer,
        string certificateThumbprint,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        // Validation lives in AppRegistrationSpec.BuildApplication; calling early surfaces argument errors
        // before we touch the Graph client.
        var newApp = AppRegistrationSpec.BuildApplication(displayName, certificateCer, certificateThumbprint);

        var client = _connection.Client
            ?? throw new InvalidOperationException("Graph no esta conectado. Usa device-code o cert primero.");
        var tenantId = _connection.TenantId
            ?? throw new InvalidOperationException("TenantId no disponible en la conexion.");

        progress?.Report(LogEntry.Info("AppReg", $"Creando App Registration '{displayName}'..."));

        var created = await client.Applications.PostAsync(newApp, cancellationToken: cancellationToken).ConfigureAwait(false)
            ?? throw new InvalidOperationException("Graph devolvio null al crear la app.");

        var appId = created.AppId ?? throw new InvalidOperationException("App creada sin AppId.");
        var objectId = created.Id ?? throw new InvalidOperationException("App creada sin Object Id.");
        progress?.Report(LogEntry.Ok("AppReg", $"App creada. AppId={appId} ObjectId={objectId}"));

        // Ensure service principal exists (Azure usually auto-creates, but explicit POST guarantees it).
        try
        {
            await client.ServicePrincipals.PostAsync(new ServicePrincipal { AppId = appId }, cancellationToken: cancellationToken).ConfigureAwait(false);
            progress?.Report(LogEntry.Ok("AppReg", "Service Principal creado."));
        }
        catch (Exception ex)
        {
            // Often returns 409 Conflict if SP already auto-provisioned. Treat as non-fatal.
            progress?.Report(LogEntry.Info("AppReg", $"Service Principal: {ex.Message}"));
        }

        var consentUrl = AppRegistrationSpec.BuildAdminConsentUrl(tenantId, appId);
        progress?.Report(LogEntry.Info("AppReg",
            $"Permisos asignados. Abre la URL de admin consent para concederlos: {consentUrl}"));

        return new AppRegistrationResult(
            AppId: appId,
            ObjectId: objectId,
            TenantId: tenantId,
            DisplayName: displayName,
            AdminConsentUrl: consentUrl);
    }
}
