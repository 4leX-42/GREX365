namespace Grex365.Core.Models;

public sealed record TransportRuleSummary(
    string Name,
    string State,
    int Priority,
    string Mode,
    string? Description);
