using Microsoft.Graph.Models;

namespace Grex365.Core.Connections;

/// <summary>
/// Pure helpers for building the App Registration payload sent to Microsoft Graph
/// when bootstrapping Grex365's service principal. Kept SDK-typed but Graph-call-free
/// so it can be unit-tested without a live tenant.
/// </summary>
public static class AppRegistrationSpec
{
    public const string GraphResourceId = "00000003-0000-0000-c000-000000000000";
    public const string ExoResourceId = "00000002-0000-0ff1-ce00-000000000000";

    public static readonly IReadOnlyList<(string Id, string Name)> GraphAppRoles = new (string Id, string Name)[]
    {
        ("741f803b-c850-494e-b5df-cde7c675a1ca", "User.ReadWrite.All"),
        ("62a82d76-70ea-41e2-9197-370581804d09", "Group.ReadWrite.All"),
        ("dbaae8cf-10b5-4b86-a4a1-f871c94c6695", "GroupMember.ReadWrite.All"),
        ("498476ce-e0fe-48b0-b801-37ba7e2685c6", "Organization.Read.All"),
        ("b0afded3-3588-46d8-8b3d-9842eff778da", "AuditLog.Read.All"),
        ("19dbc75e-c2e2-444c-a770-ec69d8559fc7", "Directory.ReadWrite.All"),
        ("230c1aed-a721-4c5d-9cb4-a90514e508ef", "Reports.Read.All"),
        ("246dd0d5-5bd0-4def-940b-0421030a5b68", "Policy.Read.All"),
        ("9a5d68dd-52b0-4cc2-bd40-abcf44ac3a30", "Application.Read.All"),
    };

    public static readonly IReadOnlyList<(string Id, string Name)> ExoAppRoles = new (string Id, string Name)[]
    {
        ("dc50a0fb-09a3-484d-be87-e023b12c6440", "Exchange.ManageAsApp"),
    };

    public static List<RequiredResourceAccess> BuildRequiredResourceAccess() => new()
    {
        new RequiredResourceAccess
        {
            ResourceAppId = GraphResourceId,
            ResourceAccess = GraphAppRoles
                .Select(r => new ResourceAccess { Id = Guid.Parse(r.Id), Type = "Role" })
                .ToList(),
        },
        new RequiredResourceAccess
        {
            ResourceAppId = ExoResourceId,
            ResourceAccess = ExoAppRoles
                .Select(r => new ResourceAccess { Id = Guid.Parse(r.Id), Type = "Role" })
                .ToList(),
        },
    };

    public static Application BuildApplication(string displayName, byte[] certificateCer, string certificateThumbprint)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(displayName);
        ArgumentNullException.ThrowIfNull(certificateCer);
        if (certificateCer.Length == 0)
        {
            throw new ArgumentException("Certificado vacío.", nameof(certificateCer));
        }
        ArgumentException.ThrowIfNullOrWhiteSpace(certificateThumbprint);

        return new Application
        {
            DisplayName = displayName,
            SignInAudience = "AzureADMyOrg",
            RequiredResourceAccess = BuildRequiredResourceAccess(),
            KeyCredentials = new List<KeyCredential>
            {
                new()
                {
                    Type = "AsymmetricX509Cert",
                    Usage = "Verify",
                    Key = certificateCer,
                    DisplayName = $"Grex365 cert ({BuildCertLabel(certificateThumbprint)})",
                },
            },
        };
    }

    public static string BuildAdminConsentUrl(string tenantId, string appId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(tenantId);
        ArgumentException.ThrowIfNullOrWhiteSpace(appId);
        return $"https://login.microsoftonline.com/{tenantId}/adminconsent?client_id={appId}";
    }

    public static string BuildCertLabel(string thumbprint) =>
        thumbprint[..Math.Min(8, thumbprint.Length)];
}
