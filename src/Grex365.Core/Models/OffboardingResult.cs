namespace Grex365.Core.Models;

public sealed record OffboardingStep(string Name, string Status, string Detail);

public sealed record OffboardingResult(
    string Upn,
    bool Success,
    IReadOnlyList<OffboardingStep> Steps);
