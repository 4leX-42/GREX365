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
}
