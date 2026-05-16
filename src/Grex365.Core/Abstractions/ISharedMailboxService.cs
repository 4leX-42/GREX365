using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface ISharedMailboxService
{
    Task<MailboxInfo?> GetMailboxAsync(string identity, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);

    Task<MailboxInfo?> ConvertToRegularAsync(string identity, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default);

    Task<MailboxPermissionResult> ApplyPermissionAsync(
        string action,
        string permission,
        string mailbox,
        string principal,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
