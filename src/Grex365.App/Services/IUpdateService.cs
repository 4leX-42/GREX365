using System.Threading;
using System.Threading.Tasks;

namespace Grex365.App.Services;

// Outcome of an update check. Status drives the Settings UI message; NewVersion is set only
// when an update is available and downloadable.
public enum UpdateCheckStatus
{
    FeedNotConfigured,   // no feed URL in preferences
    NotInstalled,        // running un-installed (dev build / portable copy) — updates don't apply
    UpToDate,
    UpdateAvailable,
    Error,
}

public sealed record UpdateCheckResult(UpdateCheckStatus Status, string? NewVersion = null, string? Detail = null);

// Velopack-backed app updates. The check is read-only; ApplyAsync downloads the pending
// update and restarts the app into the new version.
public interface IUpdateService
{
    Task<UpdateCheckResult> CheckAsync(CancellationToken cancellationToken = default);

    // Downloads the update found by the last successful CheckAsync and restarts. Throws if
    // there is no pending update.
    Task ApplyAsync(CancellationToken cancellationToken = default);
}
