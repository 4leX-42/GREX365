using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Graph.Models;

namespace Grex365.Core.Connections;

public sealed class GraphAppRegistrationService : IAppRegistrationService
{
    private const string GraphResourceId = "00000003-0000-0000-c000-000000000000";
    private const string ExoResourceId = "00000002-0000-0ff1-ce00-000000000000";

    private static readonly (string Id, string Name)[] GraphAppRoles =
    {
        ("741f803b-c850-494e-b5df-cde7c675a1ca", "User.ReadWrite.All"),
        ("62a82d76-70ea-41e2-9197-370581804d09", "Group.ReadWrite.All"),
        ("dbaae8cf-10b5-4b86-a4a1-f871c94c6695", "GroupMember.ReadWrite.All"),
        ("498476ce-e0fe-48b0-b801-37ba7e2685c6", "Organization.Read.All"),
        ("b0afded3-3588-46d8-8b3d-9842eff778da", "AuditLog.Read.All"),
        ("19dbc75e-c2e2-444c-a770-ec69d8559fc7", "Directory.ReadWrite.All"),
        ("230c1aed-a721-4c5d-9cb4-a90514e508ef", "Reports.Read.All"),
        ("246dd0d5-5bd0-4def-940b-0421030a5b68", "Policy.Read.All"),
    };

    private static readonly (string Id, string Name)[] ExoAppRoles =
    {
        ("dc50a0fb-09a3-484d-be87-e023b12c6440", "Exchange.ManageAsApp"),
    };

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
        ArgumentException.ThrowIfNullOrWhiteSpace(displayName);
        ArgumentNullException.ThrowIfNull(certificateCer);
        ArgumentException.ThrowIfNullOrWhiteSpace(certificateThumbprint);

        var client = _connection.Client
            ?? throw new InvalidOperationException("Graph no esta conectado. Usa device-code o cert primero.");
        var tenantId = _connection.TenantId
            ?? throw new InvalidOperationException("TenantId no disponible en la conexion.");

        progress?.Report(LogEntry.Info("AppReg", $"Creando App Registration '{displayName}'..."));

        var required = new List<RequiredResourceAccess>
        {
            new()
            {
                ResourceAppId = GraphResourceId,
                ResourceAccess = GraphAppRoles
                    .Select(r => new ResourceAccess { Id = Guid.Parse(r.Id), Type = "Role" })
                    .ToList(),
            },
            new()
            {
                ResourceAppId = ExoResourceId,
                ResourceAccess = ExoAppRoles
                    .Select(r => new ResourceAccess { Id = Guid.Parse(r.Id), Type = "Role" })
                    .ToList(),
            },
        };

        var newApp = new Application
        {
            DisplayName = displayName,
            SignInAudience = "AzureADMyOrg",
            RequiredResourceAccess = required,
            KeyCredentials = new List<KeyCredential>
            {
                new()
                {
                    Type = "AsymmetricX509Cert",
                    Usage = "Verify",
                    Key = certificateCer,
                    DisplayName = $"Grex365 cert ({certificateThumbprint[..Math.Min(8, certificateThumbprint.Length)]})",
                },
            },
        };

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

        var consentUrl =
            $"https://login.microsoftonline.com/{tenantId}/adminconsent?client_id={appId}";

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
