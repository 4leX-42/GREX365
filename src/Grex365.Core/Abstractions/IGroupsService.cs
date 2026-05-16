using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface IGroupsService
{
    Task<IReadOnlyList<GroupSummary>> SearchAsync(string query, CancellationToken cancellationToken = default);

    Task<IReadOnlyList<GroupMember>> GetMembersAsync(string groupId, CancellationToken cancellationToken = default);

    Task<IReadOnlyList<AddMemberResult>> AddMembersAsync(
        string groupId,
        IReadOnlyCollection<string> userIdentifiers,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task RemoveMemberAsync(
        string groupId,
        string memberId,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
