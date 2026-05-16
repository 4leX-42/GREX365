namespace Grex365.Core.Models;

public sealed record OnboardingStep(string Name, string Status, string Detail);

public sealed record OnboardingResult(
    string Upn,
    string? UserId,
    bool Success,
    IReadOnlyList<OnboardingStep> Steps);
