using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class DomainCheckViewModel : ObservableObject
{
    private readonly IDomainChecker _checker;
    private readonly IUiLogSink _log;
    private CancellationTokenSource? _cts;

    [ObservableProperty] private string _domain = string.Empty;
    [ObservableProperty] private string _statusMessage = "Indica un dominio y pulsa Comprobar.";
    [ObservableProperty] private bool _isBusy;

    public ObservableCollection<DnsRecord> Records { get; } = new();

    public DomainCheckViewModel(IDomainChecker checker, IUiLogSink log)
    {
        _checker = checker;
        _log = log;
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunAsync()
    {
        if (string.IsNullOrWhiteSpace(Domain))
        {
            StatusMessage = "Dominio vacío.";
            return;
        }
        _cts = new CancellationTokenSource();
        IsBusy = true;
        RunCommand.NotifyCanExecuteChanged();
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = $"Consultando {Domain}...";
        Records.Clear();
        try
        {
            var result = await _checker.CheckAsync(Domain.Trim(), _log.Progress, _cts.Token).ConfigureAwait(true);
            foreach (var r in result.Records)
            {
                Records.Add(r);
            }
            var problems = result.Records.Count(r => r.Status is "MISSING" or "ERROR");
            StatusMessage = problems == 0
                ? $"OK · {result.Records.Count} registros"
                : $"{problems} alertas · revisar tabla";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("DNS", ex.Message, ex));
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
