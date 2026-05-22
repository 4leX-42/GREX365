using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Win32;

namespace Grex365.App.ViewModels;

public sealed partial class AuditViewModel : ObservableObject
{
    private readonly IAuditService _audit;
    private readonly IExoForwardingAuditService _exoAudit;
    private readonly IUiLogSink _log;
    private readonly IAuditFindingsStore _findingsStore;
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

    public ObservableCollection<AuditFinding> Findings { get; } = new();
    public ICollectionView FindingsView { get; }

    public AuditViewModel(
        IAuditService audit,
        IExoForwardingAuditService exoAudit,
        IUiLogSink log,
        IAuditFindingsStore findingsStore)
    {
        _audit = audit;
        _exoAudit = exoAudit;
        _log = log;
        _findingsStore = findingsStore;
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
        StatusMessage = "Ejecutando auditoría de identidades...";
        Findings.Clear();
        try
        {
            var (summary, findings) = await _audit.RunIdentityAuditAsync(_log.Progress, _cts.Token).ConfigureAwait(true);
            Summary = summary;
            var groupFindings = await _audit.RunGroupsAuditAsync(_log.Progress, _cts.Token).ConfigureAwait(true);

            AddFindingsSorted("Identidad + grupos", findings.Concat(groupFindings));

            StatusMessage = $"{summary.UsersTotal} usuarios · {findings.Count + groupFindings.Count} hallazgos totales";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
            StatusMessage = "Umbral de inactividad debe ser >= 1.";
            return;
        }

        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = $"Descargando reporte actividad grupos (>{InactivityDays}d)...";
        Findings.Clear();
        Summary = null;
        try
        {
            var findings = await _audit
                .RunGroupActivityAuditAsync(InactivityDays, _log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("Actividad grupos", findings);
            StatusMessage = $"{findings.Count} grupos inactivos (>{InactivityDays}d).";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
        StatusMessage = "Escaneando buzones con forwarding externo...";
        Findings.Clear();
        Summary = null;
        try
        {
            var findings = await _exoAudit
                .ScanExternalForwardingAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("Forwarding externo", findings);
            StatusMessage = $"{findings.Count} forwards externos detectados.";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
        StatusMessage = "Descargando userRegistrationDetails...";
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
            StatusMessage = $"MFA: admins {summary.AdminsWithoutMfa}/{summary.AdminsTotal} sin MFA ({adminPct:F0}% cobertura) · " +
                            $"miembros {summary.MembersWithoutMfa}/{summary.MembersTotal} · " +
                            $"invitados {summary.GuestsWithoutMfa}/{summary.GuestsTotal}";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
        StatusMessage = "Descargando /oauth2PermissionGrants...";
        Findings.Clear();
        Summary = null;
        try
        {
            var (summary, findings) = await _audit
                .RunOAuthGrantsAuditAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("OAuth grants", findings);
            StatusMessage = $"OAuth grants: {summary.TotalGrants} totales · " +
                            $"{summary.UniqueClients} apps · " +
                            $"tenant-wide alto-riesgo={summary.TenantWideHighRisk} · " +
                            $"user-consented alto-riesgo={summary.UserConsentedHighRisk}";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
        StatusMessage = "Descargando /applications + credenciales...";
        Findings.Clear();
        Summary = null;
        try
        {
            var (summary, findings) = await _audit
                .RunAppCredentialsAuditAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("App credentials", findings);
            StatusMessage = $"App creds: {summary.Total} totales · " +
                            $"{summary.Expired} expired · {summary.ExpiringSoon} expiring · " +
                            $"{summary.LongLived} long-lived";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
        StatusMessage = "Descargando authorizationPolicy + securityDefaults...";
        Findings.Clear();
        Summary = null;
        try
        {
            var (summary, findings) = await _audit
                .RunTenantDefaultsAuditAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("Tenant defaults", findings);
            StatusMessage = $"Tenant defaults: SecurityDefaults={(summary.SecurityDefaultsEnabled ? "ON" : "OFF")} · " +
                            $"{findings.Count} hallazgos";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
        StatusMessage = "Enumerando roles privilegiados y miembros...";
        Findings.Clear();
        Summary = null;
        try
        {
            var (summary, findings) = await _audit
                .RunPrivilegedRolesAuditAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("Privileged roles", findings);
            StatusMessage = $"Admins: {summary.UniqueAdmins} únicos · GA={summary.GlobalAdmins} · " +
                            $"guests={summary.GuestsWithAdminRole} · disabled={summary.DisabledWithAdminRole} · " +
                            $"SP={summary.ServicePrincipalsWithAdminRole}";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
        StatusMessage = "Descargando Conditional Access policies...";
        Findings.Clear();
        Summary = null;
        try
        {
            var (summary, findings) = await _audit
                .RunConditionalAccessAuditAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("CA policies", findings);
            StatusMessage = $"CA: {summary.Total} policies · " +
                            $"{summary.Enabled} enabled · {summary.Disabled} disabled · " +
                            $"{summary.ReportOnly} report-only · {findings.Count} hallazgos";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
        StatusMessage = "Enumerando shared mailboxes + estado sign-in...";
        Findings.Clear();
        Summary = null;
        try
        {
            var (summary, findings) = await _exoAudit
                .ScanSharedMailboxSignInAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("Shared mailbox sign-in", findings);
            StatusMessage = $"Shared boxes: {summary.Total} totales · " +
                            $"sign-in enabled={summary.SignInEnabled} · " +
                            $"disabled={summary.SignInDisabled} · unknown={summary.Unknown}";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
        StatusMessage = "Get-TransportRule + Get-AcceptedDomain...";
        Findings.Clear();
        Summary = null;
        try
        {
            var (summary, findings) = await _exoAudit
                .ScanTransportRulesAsync(_log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("Transport rules", findings);
            StatusMessage = $"Transport rules: {summary.Total} totales · {summary.Enabled} enabled · " +
                            $"fwd-ext={summary.WithExternalForward} bcc-ext={summary.WithExternalBcc} " +
                            $"redir-ext={summary.WithExternalRedirect} · {findings.Count} hallazgos";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
            StatusMessage = "El tope de buzones debe ser >= 1.";
            return;
        }
        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyAllCommands();
        StatusMessage = $"Escaneando inbox rules (hasta {InboxRuleScanCap} buzones)...";
        Findings.Clear();
        Summary = null;
        try
        {
            var findings = await _exoAudit
                .ScanInboxRulesAsync(InboxRuleScanCap, _log.Progress, _cts.Token)
                .ConfigureAwait(true);
            AddFindingsSorted("Inbox rules", findings);
            StatusMessage = $"{findings.Count} reglas sospechosas detectadas.";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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

    [RelayCommand]
    private void ExportFindings()
    {
        if (Findings.Count == 0)
        {
            StatusMessage = "Sin hallazgos para exportar.";
            return;
        }

        var dlg = new SaveFileDialog
        {
            Title = "Guardar hallazgos",
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
            foreach (var f in Findings)
            {
                sb.Append(Escape(f.Category)).Append(',');
                sb.Append(Escape(f.Identity)).Append(',');
                sb.Append(Escape(f.Detail)).Append(',');
                sb.Append(Escape(f.Severity)).AppendLine();
            }
            File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            StatusMessage = $"Exportado: {Path.GetFileName(dlg.FileName)}";
            _log.Progress.Report(LogEntry.Ok("Audit", "Hallazgos exportados: " + dlg.FileName));
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Audit", ex.Message, ex));
        }
    }

    private static string Escape(string? value)
    {
        var v = value ?? string.Empty;
        if (v.Contains(',') || v.Contains('"') || v.Contains('\n'))
        {
            return '"' + v.Replace("\"", "\"\"") + '"';
        }
        return v;
    }
}
