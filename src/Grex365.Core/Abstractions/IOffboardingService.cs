using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public sealed record OffboardingOptions(
    bool DisableAccount,
    bool RemoveLicenses,
    bool ConvertMailboxToShared,
    // Dry-run: run the read-only pre-checks for real, then simulate every mutating step
    // (status "SIMULADO") without touching the tenant. Lets the flow be rehearsed safely
    // — including against production — before a real run.
    bool DryRun = false,
    // Optional mailbox finalization (EXO-only, applied after a successful conversion):
    //   ForwardTo         — SMTP forwarding to a delegate (copy kept in the shared mailbox)
    //   AutoReplyMessage  — permanent out-of-office, internal + external
    //   HideFromGal       — hide the mailbox from the global address list
    string? ForwardTo = null,
    string? AutoReplyMessage = null,
    bool HideFromGal = false,
    // Delegate that receives Full Access (+ Send As) on the resulting mailbox (via external EXO).
    string? DelegateMailboxTo = null,
    // Remove the user from all groups / distribution lists (best-effort, via Graph). Dynamic,
    // on-prem-synced and built-in groups can't be removed and are reported, not failed.
    bool RemoveFromGroups = false,
    // Enable litigation hold on the mailbox BEFORE any license removal (the hold needs the
    // EXO Plan 2 / archiving add-on entitlement present when it's applied). This is the
    // inactive-mailbox path: hold + later license removal / user deletion retains the data
    // license-free. If the hold is requested but fails, license removal is gated off.
    bool EnableLitigationHold = false,
    // Optional hold duration in days; null = indefinite (until explicitly removed).
    int? LitigationHoldDays = null);

public interface IOffboardingService
{
    // stepProgress (optional) streams each step live: a "RUNNING" report when a step starts,
    // then the final OK/ERROR/OMITIDO report — letting a panel update in real time instead
    // of waiting for the whole flow.
    Task<OffboardingResult> RunAsync(
        string upn,
        OffboardingOptions options,
        IProgress<LogEntry>? progress = null,
        IProgress<OffboardingStep>? stepProgress = null,
        CancellationToken cancellationToken = default);
}
