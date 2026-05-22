using Grex365.Core.Audit;
using Grex365.Core.Models;

namespace Grex365.Core.Abstractions;

public interface IAuditService
{
    Task<(AuditSummary Summary, IReadOnlyList<AuditFinding> Findings)> RunIdentityAuditAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<AuditFinding>> RunGroupsAuditAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<AuditFinding>> RunGroupActivityAuditAsync(
        int inactivityDays = 90,
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<(MfaCoverageSummary Summary, IReadOnlyList<AuditFinding> Findings)> RunMfaCoverageAuditAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<(CaPoliciesSummary Summary, IReadOnlyList<AuditFinding> Findings)> RunConditionalAccessAuditAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<(PrivilegedRoleSummary Summary, IReadOnlyList<AuditFinding> Findings)> RunPrivilegedRolesAuditAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<(AppCredentialsSummary Summary, IReadOnlyList<AuditFinding> Findings)> RunAppCredentialsAuditAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<(TenantDefaultsSummary Summary, IReadOnlyList<AuditFinding> Findings)> RunTenantDefaultsAuditAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);

    Task<(OAuthGrantsSummary Summary, IReadOnlyList<AuditFinding> Findings)> RunOAuthGrantsAuditAsync(
        IProgress<LogEntry>? progress = null,
        CancellationToken cancellationToken = default);
}
