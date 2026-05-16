using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public sealed record OffboardingOptions(
    bool DisableAccount,
    bool RemoveLicenses,
    bool ConvertMailboxToShared);

public interface IOffboardingService
{
    Task<OffboardingResult> RunAsync(
        string upn,
        OffboardingOptions options,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
