using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface IMailFlowRulesService
{
    Task<IReadOnlyList<TransportRuleSummary>> GetRulesAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
