using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public sealed record OnboardingOptions(
    string DisplayName,
    string Upn,
    string InitialPassword,
    string UsageLocation,
    string? MailNickname,
    IReadOnlyList<Guid> SkuIds,
    IReadOnlyList<string> GroupIdentifiers,
    bool ForceChangePasswordNextSignIn = true);

public interface IOnboardingService
{
    Task<OnboardingResult> RunAsync(
        OnboardingOptions options,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
