using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Grex365.Core.Offboarding;

namespace Grex365.App.ViewModels;

// One user queued for offboarding, with a per-account exception (delegate the resulting
// shared mailbox to someone) and a live status.
public sealed partial class OffboardingTarget : ObservableObject
{
    public OffboardingTarget(string upn, string? displayName)
    {
        Upn = upn;
        DisplayName = string.IsNullOrWhiteSpace(displayName) ? upn : displayName!;
    }

    public string Upn { get; }
    public string DisplayName { get; }

    [ObservableProperty] private string _delegateTo = string.Empty;
    // Optional per-user auto-reply override; empty = use the global template / smart default.
    [ObservableProperty] private string _autoReplyOverride = string.Empty;
    [ObservableProperty] private string _status = "PENDIENTE";
}

// One coloured line in the live log console.
public sealed record LogLine(string Time, string Text, string Level);

// An auto-reply template the user can pick from the dropdown. Body may contain {usuario} and
// {delegado} tokens, substituted per-user at run time. Empty body = free-text ("Personalizado").
public sealed record AutoReplyTemplateItem(string Name, string Body);

public sealed partial class OffboardingViewModel : ObservableObject
{
    private readonly IOffboardingService _service;
    private readonly IUiLogSink _log;
    private readonly IRbacGuard _rbac;
    private readonly IDialogService _dialogs;
    private readonly IUsersService? _users;
    private readonly IAuditService? _audit;
    private readonly ISharedMailboxService? _mailboxes;
    private readonly IClipboardService? _clipboard;
    private CancellationTokenSource? _cts;
    private CancellationTokenSource? _debounceCts;

    // Results of the last run (single or batch) — source for the CSV export.
    private readonly List<OffboardingResult> _lastResults = new();

    // --- single-user form (kept for the manual path + unit tests) ---
    [ObservableProperty] private string _upn = string.Empty;
    [ObservableProperty] private bool _disableAccount = true;
    [ObservableProperty] private bool _removeLicenses = true;
    [ObservableProperty] private bool _convertMailboxToShared = true;

    // Dry-run: rehearse the whole flow read-only (touches nothing). Safe to run in production.
    [ObservableProperty] private bool _dryRun;

    // Optional EXO finalization (applied after a successful conversion).
    [ObservableProperty] private bool _hideFromGal;
    [ObservableProperty] private bool _forwardToDelegate;
    [ObservableProperty] private string _autoReplyMessage = string.Empty;

    // Auto-reply templates: picking one fills the message box; {usuario}/{delegado} are
    // substituted per-user at run time. "Personalizado" clears to free text.
    public ObservableCollection<AutoReplyTemplateItem> AutoReplyTemplates { get; } = new();
    [ObservableProperty] private AutoReplyTemplateItem? _selectedAutoReplyTemplate;

    partial void OnSelectedAutoReplyTemplateChanged(AutoReplyTemplateItem? value)
    {
        if (value is not null) AutoReplyMessage = value.Body;
    }
    [ObservableProperty] private string _statusMessage = L10n.Get("Offboarding.Status.Initial");
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private OffboardingResult? _result;
    public ObservableCollection<OffboardingStep> Steps { get; } = new();

    // --- batch / discovery / search / live log ---
    [ObservableProperty] private string _searchQuery = string.Empty;
    [ObservableProperty] private string _delegateToAll = string.Empty;
    public ObservableCollection<UserSummary> Suggestions { get; } = new();
    public ObservableCollection<OffboardingTarget> Candidates { get; } = new();
    public ObservableCollection<OffboardingTarget> Targets { get; } = new();
    public ObservableCollection<LogLine> LiveLog { get; } = new();

    public OffboardingViewModel(
        IOffboardingService service,
        IUiLogSink log,
        IRbacGuard rbac,
        IDialogService dialogs,
        IUsersService? users = null,
        IAuditService? audit = null,
        ISharedMailboxService? mailboxes = null,
        IClipboardService? clipboard = null)
    {
        _service = service;
        _log = log;
        _rbac = rbac;
        _dialogs = dialogs;
        _users = users;
        _audit = audit;
        _mailboxes = mailboxes;
        _clipboard = clipboard;

        AutoReplyTemplates.Add(new AutoReplyTemplateItem(
            L10n.Get("Offboarding.AutoReply.Tpl.Left.Name"), L10n.Get("Offboarding.AutoReply.Tpl.Left.Body")));
        AutoReplyTemplates.Add(new AutoReplyTemplateItem(
            L10n.Get("Offboarding.AutoReply.Tpl.Contact.Name"), L10n.Get("Offboarding.AutoReply.Tpl.Contact.Body")));
        AutoReplyTemplates.Add(new AutoReplyTemplateItem(
            L10n.Get("Offboarding.AutoReply.Tpl.Custom.Name"), string.Empty));
        // Templates-first: pre-select the first so the auto-reply box starts filled.
        SelectedAutoReplyTemplate = AutoReplyTemplates[0];
    }

    // ---------- live log ----------
    private const int MaxLogLines = 600;

    private void AppendLog(string text, string level = "INFO")
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is not null && !dispatcher.CheckAccess())
        {
            dispatcher.InvokeAsync(() => AppendLog(text, level));
            return;
        }
        LiveLog.Add(new LogLine(DateTime.Now.ToString("HH:mm:ss"), text, level));
        while (LiveLog.Count > MaxLogLines) LiveLog.RemoveAt(0);
    }

    private static string LevelFromStatus(string status) => status switch
    {
        "OK" => "OK",
        "ERROR" => "ERROR",
        "OMITIDO" => "WARN",
        _ => "INFO",
    };

    [RelayCommand]
    private void ClearLog() => LiveLog.Clear();

    // Export the last run's results (single or batch) as CSV to the clipboard — an auditable
    // record to paste into a ticket / hand to HR. One row per step, with timestamps.
    [RelayCommand]
    private void ExportResults()
    {
        if (_lastResults.Count == 0)
        {
            StatusMessage = L10n.Get("Offboarding.Export.Empty");
            return;
        }
        if (_clipboard is null)
        {
            StatusMessage = L10n.Get("Offboarding.Export.NoClipboard");
            return;
        }
        var csv = OffboardingReport.ToCsv(_lastResults);
        _clipboard.SetText(csv);
        var steps = _lastResults.Sum(r => r.Steps.Count);
        StatusMessage = L10n.Format("Offboarding.Export.Copied", _lastResults.Count, steps);
        AppendLog($"Exportado CSV de {_lastResults.Count} usuario(s) al portapapeles.", "OK");
    }

    // Progress that mirrors every backend log line into the in-section live console and the
    // global log panel.
    private IProgress<LogEntry> LiveProgress() => new Progress<LogEntry>(e =>
    {
        var level = e.Severity switch
        {
            LogSeverity.Ok => "OK",
            LogSeverity.Warning => "WARN",
            LogSeverity.Error => "ERROR",
            LogSeverity.Debug => "DEBUG",
            _ => "INFO",
        };
        AppendLog($"[{e.Source}] {e.Message}", level);
        _log.Progress.Report(e);
    });

    // ---------- typeahead search (live, no Enter) ----------
    partial void OnSearchQueryChanged(string value)
    {
        if (_users is null) return;
        _debounceCts?.Cancel();
        _debounceCts = new CancellationTokenSource();
        var token = _debounceCts.Token;
        var snapshot = value ?? string.Empty;

        _ = Task.Run(async () =>
        {
            try { await Task.Delay(250, token).ConfigureAwait(false); }
            catch (OperationCanceledException) { return; }

            await Application.Current.Dispatcher.InvokeAsync(async () =>
            {
                if (token.IsCancellationRequested) return;
                if (!string.Equals(SearchQuery, snapshot, StringComparison.Ordinal)) return;
                if (snapshot.Trim().Length < 2)
                {
                    Suggestions.Clear();
                    return;
                }
                try
                {
                    var found = await _users.SearchAsync(snapshot.Trim(), token).ConfigureAwait(true);
                    if (token.IsCancellationRequested) return;
                    Suggestions.Clear();
                    foreach (var u in found.Take(15)) Suggestions.Add(u);
                }
                catch { /* surfaced elsewhere; typeahead stays quiet */ }
            });
        });
    }

    [RelayCommand]
    private void AddTarget(UserSummary? user)
    {
        if (user is null || string.IsNullOrWhiteSpace(user.UserPrincipalName)) return;
        AddTargetUpn(user.UserPrincipalName, user.DisplayName);
        SearchQuery = string.Empty;
        Suggestions.Clear();
    }

    private void AddTargetUpn(string upn, string? displayName)
    {
        if (Targets.Any(t => string.Equals(t.Upn, upn, StringComparison.OrdinalIgnoreCase))) return;
        Targets.Add(new OffboardingTarget(upn, displayName));
    }

    [RelayCommand]
    private void RemoveTarget(OffboardingTarget? target)
    {
        if (target is not null) Targets.Remove(target);
    }

    [RelayCommand]
    private void ClearTargets() => Targets.Clear();

    // ---------- candidate discovery (disabled accounts that still hold licenses) ----------
    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task FindCandidatesAsync()
    {
        if (_audit is null)
        {
            StatusMessage = L10n.Get("Offboarding.Candidates.NoService");
            return;
        }
        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyCommands();
        StatusMessage = L10n.Get("Offboarding.Candidates.Searching");
        AppendLog("Buscando usuarios (deshabilitados con licencia)…", "HEADER");
        Candidates.Clear();
        try
        {
            var (_, findings) = await _audit.RunIdentityAuditAsync(LiveProgress(), _cts.Token).ConfigureAwait(true);
            foreach (var f in findings.Where(f => f.IsAutoFixable))
            {
                Candidates.Add(new OffboardingTarget(f.Identity, f.Identity) { Status = f.Detail });
            }
            AppendLog($"{Candidates.Count} usuarios encontrados.", "HEADER");
            StatusMessage = L10n.Format("Offboarding.Candidates.Found", Candidates.Count);
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Offboarding", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyCommands();
        }
    }

    [RelayCommand]
    private void AddCandidate(OffboardingTarget? candidate)
    {
        if (candidate is not null) AddTargetUpn(candidate.Upn, candidate.DisplayName);
    }

    [RelayCommand]
    private void AddAllCandidates()
    {
        foreach (var c in Candidates) AddTargetUpn(c.Upn, c.DisplayName);
    }

    // ---------- batch run ----------
    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunBatchAsync()
    {
        if (Targets.Count == 0)
        {
            StatusMessage = L10n.Get("Offboarding.Batch.NoTargets");
            return;
        }
        if (!DisableAccount && !RemoveLicenses && !ConvertMailboxToShared)
        {
            StatusMessage = L10n.Get("Offboarding.Status.NoActionSelected");
            return;
        }

        var decision = await _rbac.EvaluateAsync().ConfigureAwait(true);
        if (!decision.Allowed)
        {
            StatusMessage = decision.Reason;
            _log.Progress.Report(LogEntry.Warn("RBAC", $"Offboarding bloqueado: {decision.Reason}"));
            return;
        }

        var ok = await _dialogs.ConfirmAsync(
            L10n.Format("Offboarding.Batch.Confirm", Targets.Count),
            L10n.Get("Offboarding.Confirm.Title"),
            DialogIcon.Warning).ConfigureAwait(true);
        if (!ok)
        {
            StatusMessage = L10n.Get("Common.Status.CancelledByUser");
            return;
        }

        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyCommands();
        var live = LiveProgress();
        var okCount = 0; var errCount = 0;
        _lastResults.Clear();
        if (DryRun) AppendLog("DRY-RUN: simulación, no se tocará el tenant.", "WARN");

        try
        {
            foreach (var target in Targets.ToList())
            {
                _cts.Token.ThrowIfCancellationRequested();
                target.Status = "RUNNING";
                AppendLog($"════ {target.Upn} ════", "HEADER");
                StatusMessage = L10n.Format("Offboarding.Status.Running", target.Upn);

                // The delegate (per-account override, else the batch default) gets FullAccess
                // below and — when "forward to delegate" is on — is also the forward target.
                var del = !string.IsNullOrWhiteSpace(target.DelegateTo) ? target.DelegateTo : DelegateToAll;
                var fwd = ForwardToDelegate && !string.IsNullOrWhiteSpace(del) ? del.Trim() : null;
                var options = new OffboardingOptions(
                    DisableAccount, RemoveLicenses, ConvertMailboxToShared, DryRun,
                    ForwardTo: fwd,
                    AutoReplyMessage: OffboardingAutoReply.Resolve(
                        AutoReplyMessage, target.AutoReplyOverride, L10n.Get("Offboarding.AutoReply.NoDelegate"),
                        target.DisplayName, del),
                    HideFromGal: HideFromGal);
                var stepProgress = new Progress<OffboardingStep>(s =>
                    AppendLog($"   [{s.Status}] {s.Name} — {s.Detail}", LevelFromStatus(s.Status)));

                try
                {
                    var result = await _service.RunAsync(target.Upn, options, live, stepProgress, _cts.Token).ConfigureAwait(true);

                    // Per-account exception: delegate the (now shared) mailbox to someone.
                    if (result.Success && !string.IsNullOrWhiteSpace(del) && _mailboxes is not null)
                    {
                        AppendLog($"   delegando buzón → {del.Trim()} (FullAccess)…", "INFO");
                        var pr = await _mailboxes.ApplyPermissionAsync("add", "FullAccess", target.Upn, del.Trim(), live, _cts.Token).ConfigureAwait(true);
                        AppendLog($"   [{pr.Status}] delegación FullAccess — {pr.Detail}", LevelFromStatus(pr.Status));
                        if (pr.Status != "OK") result = result with { Success = false };
                    }

                    _lastResults.Add(result);
                    target.Status = result.Success ? "OK" : "ERROR";
                    if (result.Success) okCount++; else errCount++;
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    target.Status = "ERROR";
                    errCount++;
                    AppendLog($"   ERROR: {ex.Message}", "ERROR");
                }
            }
            StatusMessage = L10n.Format("Offboarding.Batch.Summary", okCount, errCount, Targets.Count);
            AppendLog($"── fin: {okCount} OK · {errCount} ERROR / {Targets.Count} ──", "HEADER");
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
            AppendLog("── cancelado ──", "WARN");
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyCommands();
        }
    }

    // Entry point used by Audit "Corregir": queue the user and run the batch (live, here).
    public async Task RunCorrectiveAsync(string upn)
    {
        DisableAccount = true;
        RemoveLicenses = true;
        ConvertMailboxToShared = true;
        AddTargetUpn(upn, null);
        await RunBatchAsync().ConfigureAwait(true);
        Result = new OffboardingResult(upn, Targets.FirstOrDefault(t =>
            string.Equals(t.Upn, upn, StringComparison.OrdinalIgnoreCase))?.Status == "OK",
            Array.Empty<OffboardingStep>());
    }

    // Upserts a streamed step by Name so RUNNING→OK/ERROR updates in place (UI thread).
    private void UpsertStep(OffboardingStep step)
    {
        for (var i = 0; i < Steps.Count; i++)
        {
            if (string.Equals(Steps[i].Name, step.Name, StringComparison.Ordinal))
            {
                Steps[i] = step;
                return;
            }
        }
        Steps.Add(step);
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunAsync()
    {
        if (string.IsNullOrWhiteSpace(Upn))
        {
            StatusMessage = L10n.Get("Offboarding.Status.EmptyUpn");
            return;
        }

        var actions = new List<string>();
        if (DisableAccount) actions.Add(L10n.Get("Offboarding.Action.DisableAccount"));
        if (RemoveLicenses) actions.Add(L10n.Get("Offboarding.Action.RemoveLicenses"));
        if (ConvertMailboxToShared) actions.Add(L10n.Get("Offboarding.Action.ConvertShared"));
        if (actions.Count == 0)
        {
            StatusMessage = L10n.Get("Offboarding.Status.NoActionSelected");
            return;
        }

        var decision = await _rbac.EvaluateAsync().ConfigureAwait(true);
        if (!decision.Allowed)
        {
            StatusMessage = decision.Reason;
            _log.Progress.Report(LogEntry.Warn("RBAC", $"Offboarding bloqueado: {decision.Reason}"));
            return;
        }

        var summary = string.Join(", ", actions);
        var ok = await _dialogs.ConfirmAsync(
            L10n.Format("Offboarding.Confirm.Body", Upn, summary),
            L10n.Get("Offboarding.Confirm.Title"),
            DialogIcon.Warning).ConfigureAwait(true);
        if (!ok)
        {
            StatusMessage = L10n.Get("Common.Status.CancelledByUser");
            return;
        }

        _cts = new CancellationTokenSource();
        IsBusy = true;
        NotifyCommands();
        StatusMessage = L10n.Format("Offboarding.Status.Running", Upn);
        Steps.Clear();
        Result = null;

        try
        {
            var fwd = ForwardToDelegate && !string.IsNullOrWhiteSpace(DelegateToAll) ? DelegateToAll.Trim() : null;
            var options = new OffboardingOptions(
                DisableAccount, RemoveLicenses, ConvertMailboxToShared, DryRun,
                ForwardTo: fwd,
                AutoReplyMessage: OffboardingAutoReply.Resolve(
                    AutoReplyMessage, null, L10n.Get("Offboarding.AutoReply.NoDelegate"),
                    Upn, DelegateToAll),
                HideFromGal: HideFromGal);
            var stepProgress = new Progress<OffboardingStep>(UpsertStep);
            var result = await _service.RunAsync(Upn.Trim(), options, _log.Progress, stepProgress, _cts.Token).ConfigureAwait(true);
            Result = result;
            _lastResults.Clear();
            _lastResults.Add(result);
            if (Steps.Count == 0)
            {
                foreach (var step in result.Steps) Steps.Add(step);
            }
            StatusMessage = result.Success
                ? L10n.Format("Offboarding.Status.SuccessSummary", result.Steps.Count)
                : L10n.Format("Offboarding.Status.ErrorSummary", result.Steps.Count(s => s.Status == "ERROR"));
        }
        catch (OperationCanceledException)
        {
            StatusMessage = L10n.Get("Common.Status.Cancelled");
        }
        catch (Exception ex)
        {
            StatusMessage = L10n.Format("Common.Status.Error", ex.Message);
            _log.Progress.Report(LogEntry.Error("Offboarding", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            NotifyCommands();
        }
    }

    [RelayCommand(CanExecute = nameof(CanCancel))]
    private void Cancel() => _cts?.Cancel();

    private void NotifyCommands()
    {
        RunCommand.NotifyCanExecuteChanged();
        RunBatchCommand.NotifyCanExecuteChanged();
        FindCandidatesCommand.NotifyCanExecuteChanged();
        CancelCommand.NotifyCanExecuteChanged();
    }

    private bool CanRun() => !IsBusy;
    private bool CanCancel() => IsBusy;
}
