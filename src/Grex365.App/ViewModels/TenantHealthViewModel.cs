using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class TenantHealthViewModel : ObservableObject
{
    private readonly ITenantHealthService _service;
    private readonly IUiLogSink _log;
    private CancellationTokenSource? _cts;

    [ObservableProperty] private TenantHealth? _health;
    [ObservableProperty] private string _statusMessage = "Pulsa 'Refrescar' para cargar.";
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private int _totalConsumed;
    [ObservableProperty] private int _totalEnabled;
    [ObservableProperty] private double _overallPercent;

    public ObservableCollection<LicenseSummary> Licenses { get; } = new();

    public TenantHealthViewModel(ITenantHealthService service, IUiLogSink log)
    {
        _service = service;
        _log = log;
    }

    [RelayCommand(CanExecute = nameof(CanRefresh))]
    private async Task RefreshAsync()
    {
        _cts = new CancellationTokenSource();
        IsBusy = true;
        RefreshCommand.NotifyCanExecuteChanged();
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = "Cargando salud del tenant...";

        try
        {
            var h = await _service.GetAsync(_log.Progress, _cts.Token).ConfigureAwait(true);
            Health = h;
            Licenses.Clear();
            foreach (var l in h.Licenses.OrderByDescending(x => x.Consumed))
            {
                Licenses.Add(l);
            }
            TotalConsumed = h.Licenses.Sum(l => l.Consumed);
            TotalEnabled = h.Licenses.Sum(l => l.Enabled);
            OverallPercent = TotalEnabled > 0 ? (TotalConsumed * 100.0 / TotalEnabled) : 0;
            StatusMessage = $"{h.TotalUsers} usuarios · {h.TotalGroups} grupos · {h.Licenses.Count} SKUs · {TotalConsumed}/{TotalEnabled} asientos consumidos";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("TenantHealth", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            RefreshCommand.NotifyCanExecuteChanged();
            CancelCommand.NotifyCanExecuteChanged();
        }
    }

    [RelayCommand(CanExecute = nameof(CanCancel))]
    private void Cancel() => _cts?.Cancel();

    private bool CanRefresh() => !IsBusy;
    private bool CanCancel() => IsBusy;
}
