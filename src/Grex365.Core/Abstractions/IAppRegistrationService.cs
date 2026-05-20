using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public sealed record AppRegistrationResult(
    string AppId,
    string ObjectId,
    string TenantId,
    string DisplayName,
    string AdminConsentUrl);

public interface IAppRegistrationService
{
    Task<AppRegistrationResult> CreateAndConfigureAsync(
        string displayName,
        byte[] certificateCer,
        string certificateThumbprint,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
