using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface IUsersService
{
    Task<IReadOnlyList<UserSummary>> SearchAsync(string query, CancellationToken cancellationToken = default);

    Task<UserSummary?> GetByIdAsync(string id, CancellationToken cancellationToken = default);

    Task<IReadOnlyList<GroupSummary>> GetGroupMembershipsAsync(string userId, CancellationToken cancellationToken = default);

    Task SetAccountEnabledAsync(string userId, bool enabled, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);

    Task RemoveAllLicensesAsync(string userId, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);

    Task AssignLicenseAsync(string userId, Guid skuId, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);
}
