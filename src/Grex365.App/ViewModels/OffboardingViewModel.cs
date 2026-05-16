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
    private CancellationTokenSource? _cts;

    [ObservableProperty] private string _upn = string.Empty;
    [ObservableProperty] private bool _disableAccount = true;
    [ObservableProperty] private bool _removeLicenses = true;
    [ObservableProperty] private bool _convertMailboxToShared = true;
    [ObservableProperty] private string _statusMessage = "Indica un UPN y pulsa 'Ejecutar offboarding'.";
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private OffboardingResult? _result;

    public ObservableCollection<OffboardingStep> Steps { get; } = new();

    public OffboardingViewModel(IOffboardingService service, IUiLogSink log)
    {
        _service = service;
        _log = log;
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunAsync()
    {
        if (string.IsNullOrWhiteSpace(Upn))
        {
            StatusMessage = "UPN vacío.";
            return;
        }

        _cts = new CancellationTokenSource();
        IsBusy = true;
        RunCommand.NotifyCanExecuteChanged();
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = $"Ejecutando offboarding de {Upn}...";
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
                ? $"Offboarding OK · {result.Steps.Count} pasos"
                : $"Offboarding con errores · {result.Steps.Count(s => s.Status == "ERROR")} fallos";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
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
