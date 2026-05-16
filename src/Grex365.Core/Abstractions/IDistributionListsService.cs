using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface IDistributionListsService
{
    Task<IReadOnlyList<BulkGroupResult>> CreateFromRowsAsync(
        IReadOnlyList<BulkGroupRow> rows,
        string domain,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
