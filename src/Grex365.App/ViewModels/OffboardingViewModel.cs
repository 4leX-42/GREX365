using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class OffboardingViewModel : ObservableObject
{
    private readonly IOffboardingService _service;
    private readonly IUiLogSink _log;
    private readonly IRbacGuard _rbac;
    private readonly IDialogService _dialogs;
    private CancellationTokenSource? _cts;

    [ObservableProperty] private string _upn = string.Empty;
    [ObservableProperty] private bool _disableAccount = true;
    [ObservableProperty] private bool _removeLicenses = true;
    [ObservableProperty] private bool _convertMailboxToShared = true;
    [ObservableProperty] private string _statusMessage = L10n.Get("Offboarding.Status.Initial");
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private OffboardingResult? _result;

    public ObservableCollection<OffboardingStep> Steps { get; } = new();

    public OffboardingViewModel(IOffboardingService service, IUiLogSink log, IRbacGuard rbac, IDialogService dialogs)
    {
        _service = service;
        _log = log;
        _rbac = rbac;
        _dialogs = dialogs;
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
        RunCommand.NotifyCanExecuteChanged();
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = L10n.Format("Offboarding.Status.Running", Upn);
        Steps.Clear();
        Result = null;

        try
        {
            var options = new OffboardingOptions(DisableAccount, RemoveLicenses, ConvertMailboxToShared);
            var result = await _service.RunAsync(Upn.Trim(), options, _log.Progress, _cts.Token).ConfigureAwait(true);
            Result = result;
            foreach (var step in result.Steps)
            {
                Steps.Add(step);
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
            RunCommand.NotifyCanExecuteChanged();
            CancelCommand.NotifyCanExecuteChanged();
        }
    }

    [RelayCommand(CanExecute = nameof(CanCancel))]
    private void Cancel() => _cts?.Cancel();

    private bool CanRun() => !IsBusy;
    private bool CanCancel() => IsBusy;
}
