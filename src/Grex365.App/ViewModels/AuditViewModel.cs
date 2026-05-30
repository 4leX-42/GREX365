using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Audit;
using Grex365.Core.Models;
using Microsoft.Win32;

namespace Grex365.App.ViewModels;

public sealed partial class AuditViewModel : ObservableObject
{
    private readonly IAuditService _audit;
    private readonly IExoForwardingAuditService _exoAudit;
    private readonly IUiLogSink _log;
    private readonly IAuditFindingsStore _findingsStore;
    private readonly IGraphConnection? _graph;
    private readonly IOffboardingService? _offboarding;
    private readonly ISharedMailboxService? _mailboxes;
    private readonly IExternalExoOps? _externalExo;
    private readonly IDialogService? _dialogs;
    private readonly IRbacGuard? _rbac;
    private CancellationTokenSource? _cts;

    [ObservableProperty] private AuditSummary? _summary;
    [ObservableProperty] private string _statusMessage = "Pulsa 'Ejecutar' para auditar.";
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private int _inactivityDays = 90;
    [ObservableProperty] private int _inboxRuleScanCap = 200;

    [ObservableProperty] private bool _showErrors = true;
    [ObservableProperty] private bool _showWarnings = true;
    [ObservableProperty] private bool _showInfo = true;
    [ObservableProperty] private string _findingsFilter = string.Empty;
    [ObservableProperty] private int _errorCount;
    [ObservableProperty] private int _warningCount;
    [ObservableProperty] private int _infoCount;
    [ObservableProperty] private string? _baselineSummary;
    [ObservableProperty] private bool _hasBaseline;

    public ObservableCollection<AuditFinding> Findings { get; } = new();
    public ICollectionView FindingsView { get; }

    // Live corrective-action panel (independent side panel; not a modal dialog).
    [ObservableProperty] private bool _correctivePanelVisible;
    [ObservableProperty] private bool _correctiveRunning;
    [ObservableProperty] private string _correctiveUser = string.Empty;
    [ObservableProperty] private string _correctiveSummary = string.Empty;
    public ObservableCollection<OffboardingStep> CorrectiveSteps { get; } = new();

    // Upserts a streamed step by Name so RUNNING→OK/ERROR updates in place (UI thread).
    private void UpsertStep(OffboardingStep step)
    {
        for (var i = 0; i < CorrectiveSteps.Count; i++)
        {
            if (string.Equals(CorrectiveSteps[i].Name, step.Name, StringComparison.Ordinal))
            {
                CorrectiveSteps[i] = step;
                return;
            }
        }
        CorrectiveSteps.Add(step);
    }

    [RelayCommand]
    private void CloseCorrectivePanel() => CorrectivePanelVisible = false;

    public AuditViewModel(
        IAuditService audit,
        IExoForwardingAuditService exoAudit,
        IUiLogSink log,
        IAuditFindingsStore findingsStore,
        IGraphConnection? graph = null,
        IOffboardingService? offboarding = null,
        ISharedMailboxService? mailboxes = null,
        IDialogService? dialogs = null,
        IRbacGuard? rbac = null,
        IExternalExoOps? externalExo = null)
    {
        _audit = audit;
        _exoAudit = exoAudit;
        _log = log;
        _findingsStore = findingsStore;
        _graph = graph;
        _offboarding = offboarding;
        _mailboxes = mailboxes;
        _dialogs = dialogs;
        _rbac = rbac;
        _externalExo = externalExo;
        FindingsView = CollectionViewSource.GetDefaultView(Findings);
        FindingsView.Filter = FindingsFilterPredicate;
        Findings.CollectionChanged += (_, _) => RecomputeCounts();
    }

    private void PublishAuditResult(string auditName) =>
        _findingsStore.Update(auditName, ErrorCount, WarningCount, InfoCount);

    private bool FindingsFilterPredicate(object obj)
    {
        if (obj is not AuditFinding f)
        {
            return false;
        }
        var passesSeverity = (f.Severity ?? string.Empty).ToUpperInvariant() switch
        {
            "ERROR" => ShowErrors,
            "WARN" => ShowWarnings,
            "INFO" => ShowInfo,
            _ => true,
        };
        if (!passesSeverity)
        {
            return false;
        }
        var q = FindingsFilter?.Trim();
        if (string.IsNullOrEmpty(q))
        {
            return true;
        }
        return (f.Category?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
            || (f.Identity?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
            || (f.Detail?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false);
    }

    private void RecomputeCounts()
    {
        int err = 0, warn = 0, info = 0;
        foreach (var f in Findings)
        {
            switch ((f.Severity ?? string.Empty).ToUpperInvariant())
            {
                case "ERROR": err++; break;
                case "WARN": warn++; break;
                case "INFO": info++; break;
            }
        }
        ErrorCount = err;
        WarningCount = warn;
        InfoCount = info;
    }

    partial void OnShowErrorsChanged(bool value) => FindingsView.Refresh();
    partial void OnShowWarningsChanged(bool value) => FindingsView.Refresh();
    partial void OnShowInfoChanged(bool value) => FindingsView.Refresh();
    partial void OnFindingsFilterChanged(string value) => FindingsView.Refresh();

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunAsync()
    {
        if (IsBusy)
        {
            return;
        }
        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = L10n.Get("Audit.Status.RunningIdentity");
        Findings.Clear();
        try
        {
            // Paralelizar identity + groups audit — son endpoints distintos, no compiten.
            var identityTask = _audit.RunIdentityAuditAsync(_log.Progress, _cts.Token);
            var groupsTask = _audit.RunGroupsAuditAsync(_log.Progress, _cts.Token);
            await Task.WhenAll(identityTask, groupsTask).ConfigureAwait(true);

            var (summary, findings) = await identityTask.ConfigureAwait(true);
            var groupFindings = await groupsTask.ConfigureAwait(true);
            Summary = summary;

            AddFindingsSorted("Identidad + grupos", findings.Concat(groupFindings));

            StatusMessage = L10n.Format("Audit.Status.IdentitySummary", summary.UsersTotal, findings.Count + groupFindings.Count);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyAllCommands();
        }
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunActivityAuditAsync()
    {
        if (IsBusy)
        {
            return;
        }
        if (InactivityDays < 1)
        {
            StatusMessage = L10n.Get("Audit.Status.ThresholdMin");
            return;
        }

        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = L10n.Format("Audit.Status.DownloadingGroupActivity", InactivityDays);
        Findings.Clear();
        Summary = null;
        try
        {
            var findings = await _audit
                .RunGroupActivityAuditAsync(InactivityDays, _log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("Actividad grupos", findings);
            StatusMessage = L10n.Format("Audit.Status.InactiveGroups", findings.Count, InactivityDays);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyAllCommands();
        }
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunExternalForwardingAuditAsync()
    {
        if (IsBusy)
        {
            return;
        }
        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = L10n.Get("Audit.Status.ScanningForwarding");
        Findings.Clear();
        Summary = null;
        try
        {
            var findings = await _exoAudit
                .ScanExternalForwardingAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("Forwarding externo", findings);
            StatusMessage = L10n.Format("Audit.Status.ForwardsDetected", findings.Count);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("ExoAudit", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyAllCommands();
        }
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunMfaCoverageAuditAsync()
    {
        if (IsBusy)
        {
            return;
        }
        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = L10n.Get("Audit.Status.DownloadingUserReg");
        Findings.Clear();
        Summary = null;
        try
        {
            var (summary, findings) = await _audit
                .RunMfaCoverageAuditAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("MFA coverage", findings);
            var adminPct = summary.AdminsTotal > 0
                ? (summary.AdminsTotal - summary.AdminsWithoutMfa) * 100.0 / summary.AdminsTotal
                : 100.0;
            StatusMessage = L10n.Format("Audit.Status.MfaSummary",
                summary.AdminsWithoutMfa, summary.AdminsTotal, adminPct,
                summary.MembersWithoutMfa, summary.MembersTotal,
                summary.GuestsWithoutMfa, summary.GuestsTotal);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyAllCommands();
        }
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunOAuthGrantsAuditAsync()
    {
        if (IsBusy)
        {
            return;
        }
        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = L10n.Get("Audit.Status.DownloadingOAuth");
        Findings.Clear();
        Summary = null;
        try
        {
            var (summary, findings) = await _audit
                .RunOAuthGrantsAuditAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("OAuth grants", findings);
            StatusMessage = L10n.Format("Audit.Status.OAuthSummary",
                summary.TotalGrants, summary.UniqueClients,
                summary.TenantWideHighRisk, summary.UserConsentedHighRisk);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyAllCommands();
        }
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunAppCredentialsAuditAsync()
    {
        if (IsBusy)
        {
            return;
        }
        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = L10n.Get("Audit.Status.DownloadingApps");
        Findings.Clear();
        Summary = null;
        try
        {
            var (summary, findings) = await _audit
                .RunAppCredentialsAuditAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("App credentials", findings);
            StatusMessage = L10n.Format("Audit.Status.AppCredsSummary",
                summary.Total, summary.Expired, summary.ExpiringSoon, summary.LongLived);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyAllCommands();
        }
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunTenantDefaultsAuditAsync()
    {
        if (IsBusy)
        {
            return;
        }
        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = L10n.Get("Audit.Status.DownloadingAuthPolicy");
        Findings.Clear();
        Summary = null;
        try
        {
            var (summary, findings) = await _audit
                .RunTenantDefaultsAuditAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("Tenant defaults", findings);
            StatusMessage = L10n.Format("Audit.Status.TenantDefaultsSummary",
                summary.SecurityDefaultsEnabled ? "ON" : "OFF", findings.Count);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyAllCommands();
        }
    }

    private void AddFindingsSorted(string auditName, IEnumerable<AuditFinding> findings)
    {
        foreach (var f in findings
            .OrderBy(f => SeverityRank(f.Severity))
            .ThenBy(f => f.Category, StringComparer.OrdinalIgnoreCase)
            .ThenBy(f => f.Identity, StringComparer.OrdinalIgnoreCase))
        {
            Findings.Add(f);
        }
        PublishAuditResult(auditName);
    }

    private static int SeverityRank(string? severity) =>
        severity?.ToUpperInvariant() switch
        {
            "ERROR" => 0,
            "WARN" => 1,
            "INFO" => 2,
            _ => 3,
        };

    private void NotifyAllCommands()
    {
        RunCommand.NotifyCanExecuteChanged();
        RunActivityAuditCommand.NotifyCanExecuteChanged();
        RunExternalForwardingAuditCommand.NotifyCanExecuteChanged();
        RunInboxRulesAuditCommand.NotifyCanExecuteChanged();
        RunMfaCoverageAuditCommand.NotifyCanExecuteChanged();
        RunCaPoliciesAuditCommand.NotifyCanExecuteChanged();
        RunPrivilegedRolesAuditCommand.NotifyCanExecuteChanged();
        RunAppCredentialsAuditCommand.NotifyCanExecuteChanged();
        RunTenantDefaultsAuditCommand.NotifyCanExecuteChanged();
        RunOAuthGrantsAuditCommand.NotifyCanExecuteChanged();
        RunTransportRulesAuditCommand.NotifyCanExecuteChanged();
        RunSharedMailboxSignInAuditCommand.NotifyCanExecuteChanged();
        RunScenarioRiskyAccountsCommand.NotifyCanExecuteChanged();
        RunScenarioPrivilegedCommand.NotifyCanExecuteChanged();
        RunScenarioMailHygieneCommand.NotifyCanExecuteChanged();
        CancelCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunPrivilegedRolesAuditAsync()
    {
        if (IsBusy)
        {
            return;
        }
        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = L10n.Get("Audit.Status.EnumeratingRoles");
        Findings.Clear();
        Summary = null;
        try
        {
            var (summary, findings) = await _audit
                .RunPrivilegedRolesAuditAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("Privileged roles", findings);
            StatusMessage = L10n.Format("Audit.Status.PrivilegedSummary",
                summary.UniqueAdmins, summary.GlobalAdmins, summary.GuestsWithAdminRole,
                summary.DisabledWithAdminRole, summary.ServicePrincipalsWithAdminRole);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyAllCommands();
        }
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunCaPoliciesAuditAsync()
    {
        if (IsBusy)
        {
            return;
        }
        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = L10n.Get("Audit.Status.DownloadingCa");
        Findings.Clear();
        Summary = null;
        try
        {
            var (summary, findings) = await _audit
                .RunConditionalAccessAuditAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("CA policies", findings);
            StatusMessage = L10n.Format("Audit.Status.CaSummary",
                summary.Total, summary.Enabled, summary.Disabled, summary.ReportOnly, findings.Count);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyAllCommands();
        }
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunSharedMailboxSignInAuditAsync()
    {
        if (IsBusy)
        {
            return;
        }
        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = L10n.Get("Audit.Status.EnumeratingShared");
        Findings.Clear();
        Summary = null;
        try
        {
            var (summary, findings) = await _exoAudit
                .ScanSharedMailboxSignInAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("Shared mailbox sign-in", findings);
            StatusMessage = L10n.Format("Audit.Status.SharedSummary",
                summary.Total, summary.SignInEnabled, summary.SignInDisabled, summary.Unknown);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("ExoAudit", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyAllCommands();
        }
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunTransportRulesAuditAsync()
    {
        if (IsBusy)
        {
            return;
        }
        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = L10n.Get("Audit.Status.TransportScan");
        Findings.Clear();
        Summary = null;
        try
        {
            var (summary, findings) = await _exoAudit
                .ScanTransportRulesAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("Transport rules", findings);
            StatusMessage = L10n.Format("Audit.Status.TransportSummary",
                summary.Total, summary.Enabled, summary.WithExternalForward,
                summary.WithExternalBcc, summary.WithExternalRedirect, findings.Count);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("ExoAudit", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyAllCommands();
        }
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunInboxRulesAuditAsync()
    {
        if (IsBusy)
        {
            return;
        }
        if (InboxRuleScanCap < 1)
        {
            StatusMessage = L10n.Get("Audit.Status.MailboxCapMin");
            return;
        }
        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = L10n.Format("Audit.Status.ScanningInbox", InboxRuleScanCap);
        Findings.Clear();
        Summary = null;
        try
        {
            var findings = await _exoAudit
                .ScanInboxRulesAsync(InboxRuleScanCap, _log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("Inbox rules", findings);
            StatusMessage = L10n.Format("Audit.Status.SuspiciousRules", findings.Count);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("ExoAudit", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyAllCommands();
        }
    }

    [RelayCommand(CanExecute = nameof(CanCancel))]
    private void Cancel() => _cts?.Cancel();

    private bool CanRun() => !IsBusy;
    private bool CanCancel() => IsBusy;

    // Guided corrective flow for a "Disabled+License" finding: block sign-in + revoke
    // sessions → convert mailbox to shared (preserves mail) → release licenses. Follows
    // Microsoft's recommended order; the account stays disabled as the shared-mailbox anchor.
    [RelayCommand]
    private async Task FixDisabledWithLicenseAsync(AuditFinding? finding)
    {
        if (finding is null || !finding.IsAutoFixable)
        {
            StatusMessage = L10n.Get("Audit.Fix.NotApplicable");
            return;
        }
        if (_offboarding is null || _dialogs is null)
        {
            return;
        }
        if (IsBusy)
        {
            return;
        }

        var upn = finding.Identity;
        if (string.IsNullOrWhiteSpace(upn) || upn.StartsWith('('))
        {
            StatusMessage = L10n.Get("Audit.Fix.NotApplicable");
            return;
        }

        if (_rbac is not null)
        {
            var decision = await _rbac.EvaluateAsync().ConfigureAwait(true);
            if (!decision.Allowed)
            {
                StatusMessage = decision.Reason;
                _log.Progress.Report(LogEntry.Warn("RBAC", $"Corrección offboarding bloqueada: {decision.Reason}"));
                return;
            }
        }

        // Quick confirmation (the detailed plan + warnings now stream live in the panel).
        var ok = await _dialogs.ConfirmAsync(
            L10n.Format("Audit.Fix.Body", upn, string.Empty),
            L10n.Get("Audit.Fix.Title"),
            DialogIcon.Warning).ConfigureAwait(true);
        if (!ok)
        {
            StatusMessage = L10n.Get("Common.Status.CancelledByUser");
            return;
        }

        // Open the independent live panel.
        CorrectiveSteps.Clear();
        CorrectiveUser = upn;
        CorrectiveSummary = string.Empty;
        CorrectiveRunning = true;
        CorrectivePanelVisible = true;

        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = L10n.Format("Audit.Fix.Running", upn);

        // Progress captured on the UI thread → callbacks marshal back here, safe for the
        // ObservableCollection bound to the panel.
        var stepProgress = new Progress<OffboardingStep>(UpsertStep);

        try
        {
            // Step 1 (visible): pre-check the mailbox so we surface holds/size/already-shared.
            UpsertStep(new OffboardingStep("Pre-check buzón", "RUNNING", "…"));
            MailboxInfo? mailbox = null;
            try
            {
                if (_externalExo is not null)
                {
                    mailbox = await _externalExo.GetMailboxFactsAsync(upn, _log.Progress, _cts.Token).ConfigureAwait(true);
                }
                else if (_mailboxes is not null)
                {
                    mailbox = await _mailboxes.GetMailboxAsync(upn, _log.Progress).ConfigureAwait(true);
                }
            }
            catch (Exception ex)
            {
                _log.Progress.Report(LogEntry.Info("Audit", $"Pre-check buzón {upn} no disponible: {ex.Message}"));
            }

            var alreadyShared = mailbox?.IsSharedMailbox == true;
            if (mailbox is null)
            {
                UpsertStep(new OffboardingStep("Pre-check buzón", "OMITIDO", "No se pudo leer el buzón (¿Exchange?). Continuando; si la conversión falla NO se quitan licencias."));
            }
            else
            {
                var facts = $"Tipo={mailbox.RecipientTypeDetails}";
                if (mailbox.TotalItemSizeGb is { } gb) facts += $" · {gb} GB";
                if (mailbox.HasBlockingHold) facts += " · ⚠ retención (hold)";
                if (mailbox.ExceedsUnlicensedSharedLimit) facts += " · ⚠ >50 GB";
                UpsertStep(new OffboardingStep("Pre-check buzón", "OK", facts));
            }

            var options = new OffboardingOptions(
                DisableAccount: true,
                RemoveLicenses: true,
                ConvertMailboxToShared: !alreadyShared);

            var result = await _offboarding.RunAsync(upn, options, _log.Progress, stepProgress, _cts.Token).ConfigureAwait(true);

            var okSteps = result.Steps.Count(s => s.Status == "OK");
            CorrectiveSummary = result.Success
                ? L10n.Format("Audit.Fix.Done", upn, okSteps, result.Steps.Count)
                : L10n.Format("Audit.Fix.Failed", upn);
            StatusMessage = CorrectiveSummary;

            if (result.Success)
            {
                var stale = Findings.FirstOrDefault(f => f.IsAutoFixable &&
                    string.Equals(f.Identity, upn, StringComparison.OrdinalIgnoreCase));
                if (stale is not null)
                {
                    Findings.Remove(stale);
                }
            }
        }
        catch (OperationCanceledException)
        {
            CorrectiveSummary = L10n.Get("Common.Status.Cancelled");
            StatusMessage = CorrectiveSummary;
        }
        catch (Exception ex)
        {
            CorrectiveSummary = L10n.Format("Common.Status.Error", ex.Message);
            StatusMessage = CorrectiveSummary;
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
        finally
        {
            CorrectiveRunning = false;
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyAllCommands();
        }
    }

    // ---- Pre-canned scenarios (combos) -----------------------------------

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunScenarioRiskyAccountsAsync()
    {
        await RunScenarioCommonAsync(
            "Cuentas en riesgo",
            "Identidad · deshab+licencia / cuentas inactivas...",
            async () =>
            {
                // Only the identity audit matters here — the keep-filter discards every
                // group finding, so running the (slow, per-group) groups audit was pure
                // wasted work. Identity-only makes this scenario much faster.
                var (summary, findings) = await _audit
                    .RunIdentityAuditAsync(_log.Progress, _cts!.Token)
                    .ConfigureAwait(true);
                Summary = summary;
                return findings;
            },
            keep: f =>
                f.Category.Contains("disabled", StringComparison.OrdinalIgnoreCase) ||
                f.Category.Contains("license", StringComparison.OrdinalIgnoreCase) ||
                f.Category.Contains("stale", StringComparison.OrdinalIgnoreCase) ||
                f.Category.Contains("inactive", StringComparison.OrdinalIgnoreCase)).ConfigureAwait(true);
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunScenarioPrivilegedAsync()
    {
        await RunScenarioCommonAsync(
            "Acceso privilegiado",
            "Privileged roles + MFA + CA policies en paralelo...",
            async () =>
            {
                var rolesTask = _audit.RunPrivilegedRolesAuditAsync(_log.Progress, _cts!.Token);
                var mfaTask = _audit.RunMfaCoverageAuditAsync(_log.Progress, _cts!.Token);
                var caTask = _audit.RunConditionalAccessAuditAsync(_log.Progress, _cts!.Token);
                await Task.WhenAll(rolesTask, mfaTask, caTask).ConfigureAwait(true);
                var rolesResult = await rolesTask.ConfigureAwait(true);
                var mfaResult = await mfaTask.ConfigureAwait(true);
                var caResult = await caTask.ConfigureAwait(true);
                return rolesResult.Findings
                    .Concat(mfaResult.Findings)
                    .Concat(caResult.Findings);
            }).ConfigureAwait(true);
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunScenarioMailHygieneAsync()
    {
        await RunScenarioCommonAsync(
            "Higiene mail (BEC)",
            "Forwarding externo + Inbox rules + Transport rules + Shared mailbox sign-in...",
            async () =>
            {
                var fwd = await _exoAudit.ScanExternalForwardingAsync(_log.Progress, _cts!.Token).ConfigureAwait(true);
                var rules = await _exoAudit.ScanInboxRulesAsync(InboxRuleScanCap, _log.Progress, _cts!.Token).ConfigureAwait(true);
                var tr = await _exoAudit.ScanTransportRulesAsync(_log.Progress, _cts!.Token).ConfigureAwait(true);
                var sh = await _exoAudit.ScanSharedMailboxSignInAsync(_log.Progress, _cts!.Token).ConfigureAwait(true);
                return fwd
                    .Concat(rules)
                    .Concat(tr.Findings)
                    .Concat(sh.Findings);
            }).ConfigureAwait(true);
    }

    private async Task RunScenarioCommonAsync(
        string name,
        string startMessage,
        Func<Task<IEnumerable<AuditFinding>>> runner,
        Func<AuditFinding, bool>? keep = null)
    {
        if (IsBusy) return;
        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = startMessage;
        Findings.Clear();
        Summary = null;
        try
        {
            var raw = await runner().ConfigureAwait(true);
            var filtered = keep is null ? raw : raw.Where(keep);
            AddFindingsSorted(name, filtered);
            StatusMessage = L10n.Format("Audit.Status.ScenarioSummary", name, Findings.Count);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyAllCommands();
        }
    }

    [RelayCommand]
    private void ExportFindings()
    {
        if (Findings.Count == 0)
        {
            StatusMessage = L10n.Get("Audit.Status.NoFindingsToExport");
            return;
        }

        var dlg = new SaveFileDialog
        {
            Title = L10n.Get("Audit.Dialog.SaveFindings"),
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"identity_audit_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };
        if (dlg.ShowDialog() != true)
        {
            return;
        }

        try
        {
            var sb = new StringBuilder();
            sb.AppendLine("Categoria,Identity,Detalle,Severidad");
            var rows = FindingsView.Cast<AuditFinding>().ToList();
            foreach (var f in rows)
            {
                sb.Append(Escape(f.Category)).Append(',');
                sb.Append(Escape(f.Identity)).Append(',');
                sb.Append(Escape(f.Detail)).Append(',');
                sb.Append(Escape(f.Severity)).AppendLine();
            }
            File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            StatusMessage = L10n.Format("Common.Status.Exported", Path.GetFileName(dlg.FileName));
            _log.Progress.Report(LogEntry.Ok("Audit", "Hallazgos exportados: " + dlg.FileName));
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
    }

    private static string Escape(string? value) => Grex365.Core.Csv.CsvEscaper.Escape(value);

    [RelayCommand]
    private void LoadBaseline()
    {
        var dlg = new OpenFileDialog
        {
            Title = L10n.Get("Audit.Dialog.LoadBaseline"),
            Filter = "JSON (*.json)|*.json|Todos|*.*"
        };
        if (dlg.ShowDialog() != true)
        {
            return;
        }

        try
        {
            var json = File.ReadAllText(dlg.FileName);
            var envelope = AuditReportJsonBuilder.Parse(json);
            if (envelope is null)
            {
                StatusMessage = L10n.Get("Audit.Status.BaselineEmpty");
                return;
            }

            var current = FindingsView.Cast<AuditFinding>().ToList();
            var diff = AuditBaselineComparer.Compare(envelope.Findings, current);
            BaselineSummary =
                $"Baseline {Path.GetFileName(dlg.FileName)} ({envelope.GeneratedAt:yyyy-MM-dd}) — " +
                $"nuevos: {diff.NewCount} · resueltos: {diff.ResolvedCount} · persistentes: {diff.PersistentCount}";
            HasBaseline = true;
            StatusMessage = BaselineSummary;
            _log.Progress.Report(LogEntry.Ok("Audit",
                $"Baseline cargado: {diff.NewCount} nuevos, {diff.ResolvedCount} resueltos, {diff.PersistentCount} persistentes"));
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Audit.Status.BaselineError", ex.Message);
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
    }

    [RelayCommand]
    private void ClearBaseline()
    {
        BaselineSummary = null;
        HasBaseline = false;
        StatusMessage = L10n.Get("Audit.Status.BaselineCleared");
    }

    [RelayCommand]
    private void ExportFindingsJson()
    {
        if (Findings.Count == 0)
        {
            StatusMessage = L10n.Get("Audit.Status.NoFindingsToExport");
            return;
        }

        var dlg = new SaveFileDialog
        {
            Title = L10n.Get("Audit.Dialog.SaveJsonReport"),
            Filter = "JSON (*.json)|*.json",
            FileName = $"audit_report_{DateTime.Now:yyyyMMdd_HHmmss}.json"
        };
        if (dlg.ShowDialog() != true)
        {
            return;
        }

        try
        {
            var rows = FindingsView.Cast<AuditFinding>().ToList();
            var ctx = new AuditReportContext(
                Title: L10n.Get("Audit.Report.Title"),
                GeneratedAt: DateTime.Now,
                TenantDomain: _graph?.TenantId,
                GeneratedBy: Environment.UserName);
            var json = AuditReportJsonBuilder.Build(rows, ctx);
            File.WriteAllText(dlg.FileName, json, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            StatusMessage = L10n.Format("Common.Status.Exported", Path.GetFileName(dlg.FileName));
            _log.Progress.Report(LogEntry.Ok("Audit", "Informe JSON exportado: " + dlg.FileName));
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
    }

    [RelayCommand]
    private void ExportFindingsHtml()
    {
        if (Findings.Count == 0)
        {
            StatusMessage = L10n.Get("Audit.Status.NoFindingsToExport");
            return;
        }

        var dlg = new SaveFileDialog
        {
            Title = L10n.Get("Audit.Dialog.SaveHtmlReport"),
            Filter = "HTML (*.html)|*.html",
            FileName = $"audit_report_{DateTime.Now:yyyyMMdd_HHmmss}.html"
        };
        if (dlg.ShowDialog() != true)
        {
            return;
        }

        try
        {
            var rows = FindingsView.Cast<AuditFinding>().ToList();
            var ctx = new AuditReportContext(
                Title: L10n.Get("Audit.Report.Title"),
                GeneratedAt: DateTime.Now,
                TenantDomain: _graph?.TenantId,
                GeneratedBy: Environment.UserName);
            var html = AuditReportHtmlBuilder.Build(rows, ctx);
            File.WriteAllText(dlg.FileName, html, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            StatusMessage = L10n.Format("Common.Status.Exported", Path.GetFileName(dlg.FileName));
            _log.Progress.Report(LogEntry.Ok("Audit", "Informe HTML exportado: " + dlg.FileName));
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
    }
}
