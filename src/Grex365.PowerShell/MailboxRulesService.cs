using Grex365.Core.Abstractions;
using Grex365.Core.Mailboxes;
using Grex365.Core.Models;

namespace Grex365.PowerShell;

// Thin wrapper over IExternalExoOps: auto-reply, forwarding and calendar permissions all run in
// an external pwsh process. The in-proc RunspacePool path is unreliable for EXO V3 (the
// "HttpResponseMessage does not contain GetResponseHeader" failure), so every cmdlet delegates to
// the external host. Input validation (pure, testable) stays here and short-circuits before EXO.
public sealed class MailboxRulesService : IMailboxRulesService
{
    private readonly IExternalExoOps _exo;

    public MailboxRulesService(IExternalExoOps exo)
    {
        _exo = exo;
    }

    public Task<AutoReplyConfig?> GetAutoReplyAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default) =>
        _exo.GetAutoReplyAsync(identity, progress, cancellationToken);

    public async Task SetAutoReplyAsync(
        string identity,
        AutoReplyConfig config,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var errors = MailboxRulesValidator.ValidateAutoReply(config);
        if (errors.Count > 0)
        {
            throw new ArgumentException("AutoReply inválido: " + string.Join("; ", errors));
        }
        await _exo.SetAutoReplyConfigAsync(identity, config, progress, cancellationToken).ConfigureAwait(false);
        progress?.Report(LogEntry.Ok("Mailbox", $"AutoReply {config.State} en {identity}"));
    }

    public Task<ForwardingConfig?> GetForwardingAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default) =>
        _exo.GetForwardingAsync(identity, progress, cancellationToken);

    public async Task SetForwardingAsync(
        string identity,
        string forwardingSmtpAddress,
        bool deliverToMailboxAndForward,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var errors = MailboxRulesValidator.ValidateForwarding(forwardingSmtpAddress);
        if (errors.Count > 0)
        {
            throw new ArgumentException("Forwarding inválido: " + string.Join("; ", errors));
        }
        await _exo.ConfigureForwardingAsync(identity, forwardingSmtpAddress, deliverToMailboxAndForward, progress, cancellationToken).ConfigureAwait(false);
        progress?.Report(LogEntry.Ok("Mailbox", $"Forwarding {identity} → {forwardingSmtpAddress}"));
    }

    public async Task ClearForwardingAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        await _exo.ClearForwardingAsync(identity, progress, cancellationToken).ConfigureAwait(false);
        progress?.Report(LogEntry.Ok("Mailbox", $"Forwarding limpiado en {identity}"));
    }

    public Task<IReadOnlyList<CalendarPermissionEntry>> GetCalendarPermissionsAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default) =>
        _exo.GetCalendarPermissionsAsync(identity, progress, cancellationToken);

    public Task ApplyCalendarPermissionAsync(
        string identity,
        string principal,
        string accessRights,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(identity) || string.IsNullOrWhiteSpace(principal))
        {
            throw new ArgumentException("Identity y Principal requeridos.");
        }
        if (string.IsNullOrWhiteSpace(accessRights) || !CalendarAccessRights.All.Contains(accessRights))
        {
            throw new ArgumentException("AccessRights inválido: " + accessRights);
        }
        return _exo.ApplyCalendarPermissionAsync(identity, principal, accessRights, progress, cancellationToken);
    }

    public Task RemoveCalendarPermissionAsync(
        string identity,
        string principal,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(identity) || string.IsNullOrWhiteSpace(principal))
        {
            throw new ArgumentException("Identity y Principal requeridos.");
        }
        return _exo.RemoveCalendarPermissionAsync(identity, principal, progress, cancellationToken);
    }
}
