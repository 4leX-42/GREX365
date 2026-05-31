namespace Grex365.Core.Models;

// At = wall-clock time the step finished (null while RUNNING / for synthetic steps).
public sealed record OffboardingStep(string Name, string Status, string Detail, DateTimeOffset? At = null);

public sealed record OffboardingResult(
    string Upn,
    bool Success,
    IReadOnlyList<OffboardingStep> Steps,
    bool DryRun = false,
    DateTimeOffset? StartedAt = null,
    DateTimeOffset? EndedAt = null);
