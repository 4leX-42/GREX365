using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface IDomainChecker
{
    Task<DomainCheck> CheckAsync(string domain, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);
}
