using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

// Runs Exchange Online operations in an external pwsh.exe process (connect-per-call with
// the app certificate). The in-process RunspacePool path is unreliable for EXO V3: the
// session lives in a single pooled runspace (others land session-less) and the REST/RPS
// toggle trips the "HttpResponseMessage does not contain GetResponseHeader" error. An
// external full PowerShell host runs the modern REST cmdlets correctly.
public interface IExternalExoOps
{
    // Best-effort pre-check: mailbox type, holds, archive, size. Null if it can't be read.
    Task<MailboxInfo?> GetMailboxFactsAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Converts the mailbox to shared (Set-Mailbox -Type Shared) and waits for propagation.
    // Returns the final mailbox state.
    Task<MailboxInfo?> ConvertToSharedAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Enables a permanent auto-reply (out-of-office) with the same message internal + external.
    Task SetAutoReplyAsync(
        string identity,
        string message,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Sets SMTP forwarding to a delegate, keeping a copy in the (now shared) mailbox.
    Task SetForwardingAsync(
        string identity,
        string forwardTo,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Hides the mailbox from address lists (GAL). Returns a short note; may report a SKIP for
    // hybrid objects synced from on-prem AD (which must be hidden in local AD instead).
    Task<string> HideFromGalAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Enables litigation hold on the mailbox (Set-Mailbox -LitigationHoldEnabled $true), with an
    // optional duration in days (null = indefinite). Idempotent: already-on-hold reports a note
    // instead of failing. Requires Exchange Online Plan 2 / archiving add-on on the mailbox.
    Task<string> SetLitigationHoldAsync(
        string identity,
        int? durationDays = null,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Grants a delegate Full Access (AutoMapping off) and, optionally, Send As on the mailbox.
    // Runs via external pwsh — the in-proc EXO path trips the GetResponseHeader bug. Returns a note.
    Task<string> GrantDelegateAsync(
        string mailbox,
        string delegateUpn,
        bool sendAs,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Converts the mailbox back to a regular user mailbox (Set-Mailbox -Type Regular) and waits.
    Task<MailboxInfo?> ConvertToRegularAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Adds/removes a mailbox permission (FullAccess / SendAs / SendOnBehalf). Inputs are assumed
    // pre-validated by the caller; returns OK/ERROR.
    Task<MailboxPermissionResult> ApplyPermissionAsync(
        string action,
        string permission,
        string mailbox,
        string principal,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Lists the explicit FullAccess / SendAs / SendOnBehalf delegates on the mailbox.
    Task<IReadOnlyList<MailboxPermissionEntry>> GetPermissionsAsync(
        string mailbox,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // ---- Mailbox rules (OOO / forwarding / calendar) — used by MailboxRulesService ----
    // These ran in-proc and tripped the GetResponseHeader bug; they go through external pwsh now.

    // Reads the auto-reply (out-of-office) configuration. Null if the mailbox can't be read.
    Task<AutoReplyConfig?> GetAutoReplyAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Sets the full auto-reply configuration (state + internal/external message + optional window).
    // Unlike the offboarding SetAutoReplyAsync, this honours every field of the config as-is.
    Task SetAutoReplyConfigAsync(
        string identity,
        AutoReplyConfig config,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Reads the current SMTP/internal forwarding configuration. Null if the mailbox can't be read.
    Task<ForwardingConfig?> GetForwardingAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Sets SMTP forwarding with an explicit "keep a copy" flag (parameterised, unlike the
    // offboarding SetForwardingAsync which always keeps a copy).
    Task ConfigureForwardingAsync(
        string identity,
        string forwardingSmtpAddress,
        bool deliverToMailboxAndForward,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Clears any SMTP/internal forwarding and the deliver-and-forward flag.
    Task ClearForwardingAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Lists the non-default calendar folder permissions on the mailbox.
    Task<IReadOnlyList<CalendarPermissionEntry>> GetCalendarPermissionsAsync(
        string identity,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Adds or updates a calendar folder permission for a principal (Add- or Set-, as needed).
    Task ApplyCalendarPermissionAsync(
        string identity,
        string principal,
        string accessRights,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Removes a principal's calendar folder permission.
    Task RemoveCalendarPermissionAsync(
        string identity,
        string principal,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    // Removes a member from classic distribution lists / mail-enabled security groups, whose
    // membership lives in Exchange Online (Graph can't write it). One pwsh invocation: connects
    // once, iterates the groups, returns one per-group result (best-effort — a failing group
    // doesn't stop the rest).
    Task<IReadOnlyList<DistributionGroupRemovalResult>> RemoveFromDistributionGroupsAsync(
        string memberIdentity,
        IReadOnlyList<string> groupIdentities,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
