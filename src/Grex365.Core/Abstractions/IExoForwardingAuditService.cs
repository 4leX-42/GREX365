using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface IExoForwardingAuditService
{
    Task<IReadOnlyList<AuditFinding>> ScanExternalForwardingAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
