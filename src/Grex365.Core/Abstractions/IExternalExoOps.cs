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

    // Grants a delegate Full Access (AutoMapping off) and, optionally, Send As on the mailbox.
    // Runs via external pwsh — the in-proc EXO path trips the GetResponseHeader bug. Returns a note.
    Task<string> GrantDelegateAsync(
        string mailbox,
        string delegateUpn,
        bool sendAs,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
