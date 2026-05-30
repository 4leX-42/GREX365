using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public sealed record OffboardingOptions(
    bool DisableAccount,
    bool RemoveLicenses,
    bool ConvertMailboxToShared);

public interface IOffboardingService
{
    // stepProgress (optional) streams each step live: a "RUNNING" report when a step starts,
    // then the final OK/ERROR/OMITIDO report — letting a panel update in real time instead
    // of waiting for the whole flow.
    Task<OffboardingResult> RunAsync(
        string upn,
        OffboardingOptions options,
        IProgress<LogEntry>? progress = null,
        IProgress<OffboardingStep>? stepProgress = null,
        CancellationToken cancellationToken = default);
}
