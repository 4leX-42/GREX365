using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface IExoForwardingAuditService
{
    Task<IReadOnlyList<AuditFinding>> ScanExternalForwardingAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<AuditFinding>> ScanInboxRulesAsync(
        int maxMailboxes = 200,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
