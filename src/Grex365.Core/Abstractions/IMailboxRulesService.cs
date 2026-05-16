using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface IMailboxRulesService
{
    Task<AutoReplyConfig?> GetAutoReplyAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task SetAutoReplyAsync(
        string identity,
        AutoReplyConfig config,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<ForwardingConfig?> GetForwardingAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task SetForwardingAsync(
        string identity,
        string forwardingSmtpAddress,
        bool deliverToMailboxAndForward,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task ClearForwardingAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<CalendarPermissionEntry>> GetCalendarPermissionsAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task ApplyCalendarPermissionAsync(
        string identity,
        string principal,
        string accessRights,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task RemoveCalendarPermissionAsync(
        string identity,
        string principal,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
