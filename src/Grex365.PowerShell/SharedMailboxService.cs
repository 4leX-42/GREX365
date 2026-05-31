using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

// Thin wrapper over IExternalExoOps: all Exchange Online work runs in an external pwsh process.
// The in-proc RunspacePool path is unreliable for EXO V3 (the "HttpResponseMessage does not
// contain GetResponseHeader" failure), so lookups / conversions / permissions all delegate to
// the external host. Input validation for permission changes stays here (pure, testable).
public sealed class SharedMailboxService : ISharedMailboxService
{
    private readonly IExternalExoOps _exo;

    public SharedMailboxService(IExternalExoOps exo)
    {
        _exo = exo;
    }

    public Task<MailboxInfo?> GetMailboxAsync(string identity, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default) =>
        _exo.GetMailboxFactsAsync(identity, progress, cancellationToken);

    public Task<MailboxInfo?> ConvertToSharedAsync(string identity, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default) =>
        _exo.ConvertToSharedAsync(identity, progress, cancellationToken);

    public Task<MailboxInfo?> ConvertToRegularAsync(string identity, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default) =>
        _exo.ConvertToRegularAsync(identity, progress, cancellationToken);

    public async Task<MailboxPermissionResult> ApplyPermissionAsync(
        string action,
        string permission,
        string mailbox,
        string principal,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var a = (action ?? string.Empty).Trim();
        var p = (permission ?? string.Empty).Trim();
        var m = mailbox ?? string.Empty;
        var pr = principal ?? string.Empty;
        var actionLower = a.ToLowerInvariant();

        if (actionLower != "add" && actionLower != "remove")
        {
            return new MailboxPermissionResult(a, p, m, pr, "INVALIDO", "Action debe ser add|remove");
        }
        if (p is not ("FullAccess" or "SendAs" or "SendOnBehalf"))
        {
            return new MailboxPermissionResult(a, p, m, pr, "INVALIDO", "Permission no soportada");
        }
        if (string.IsNullOrWhiteSpace(m) || string.IsNullOrWhiteSpace(pr))
        {
            return new MailboxPermissionResult(a, p, m, pr, "INVALIDO", "Mailbox/Principal vacío");
        }

        return await _exo.ApplyPermissionAsync(actionLower, p, m, pr, progress, cancellationToken).ConfigureAwait(false);
    }

    public Task<IReadOnlyList<MailboxPermissionEntry>> GetPermissionsAsync(string mailbox, IProgress<LogEntry>? progress = null, CancellationToken cancellationToken = default) =>
        _exo.GetPermissionsAsync(mailbox, progress, cancellationToken);
}
