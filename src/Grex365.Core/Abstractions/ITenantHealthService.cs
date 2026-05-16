using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface ITenantHealthService
{
    Task<TenantHealth> GetAsync(IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);
}
