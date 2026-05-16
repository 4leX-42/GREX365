using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public sealed record AuditRecord(
    DateTimeOffset Timestamp,
    string Actor,
    string Source,
    string Outcome,
    string Message,
    string? Detail = null);

public interface IAuditLog
{
    Task WriteAsync(AuditRecord record, CancellationToken cancellationToken = default);

    Task<IReadOnlyList<AuditRecord>> ReadMonthAsync(int year, int month, CancellationToken cancellationToken = default);

    string GetMonthFilePath(int year, int month);
}
