using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface IUsersService
{
    Task<IReadOnlyList<UserSummary>> SearchAsync(string query, CancellationToken cancellationToken = default);

    Task<UserSummary?> GetByIdAsync(string id, CancellationToken cancellationToken = default);

    Task<IReadOnlyList<GroupSummary>> GetGroupMembershipsAsync(string userId, CancellationToken cancellationToken = default);

    Task SetAccountEnabledAsync(string userId, bool enabled, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);

    Task RemoveAllLicensesAsync(string userId, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);

    Task AssignLicenseAsync(string userId, Guid skuId, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);

    Task<IReadOnlyList<SkuInfo>> ListSkusAsync(CancellationToken cancellationToken = default);

    Task<IReadOnlyList<Guid>> GetAssignedLicensesAsync(string userId, CancellationToken cancellationToken = default);

    Task RemoveLicenseAsync(string userId, Guid skuId, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);

    Task<string> ResetPasswordAsync(string userId, bool forceChangeNextSignIn = true, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);

    Task RevokeSignInSessionsAsync(string userId, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);

    // ACTIVE Entra directory roles the user holds (memberOf — PIM-eligible don't surface).
    Task<IReadOnlyList<DirectoryRoleSummary>> GetDirectoryRolesAsync(string userId, CancellationToken cancellationToken = default);

    // Removes the user from one directory role. Requires RoleManagement.ReadWrite.Directory.
    Task RemoveFromDirectoryRoleAsync(string roleId, string userId, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);

    // Registered authentication methods (MFA). Requires UserAuthenticationMethod.Read.All.
    Task<IReadOnlyList<AuthMethodSummary>> GetAuthMethodsAsync(string userId, CancellationToken cancellationToken = default);

    // Deletes one auth method (must be Removable). Requires UserAuthenticationMethod.ReadWrite.All.
    Task RemoveAuthMethodAsync(string userId, AuthMethodSummary method, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);

    // Sends a plain-text mail from the given mailbox (app-only sendMail). Requires Mail.Send.
    Task SendMailAsync(string fromUserIdOrUpn, string to, string subject, string body, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);

    Task<UserSummary> CreateUserAsync(
        NewUserSpec spec,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}

public sealed record NewUserSpec(
    string DisplayName,
    string UserPrincipalName,
    string MailNickname,
    string Password,
    string UsageLocation,
    bool ForceChangePasswordNextSignIn = true);
