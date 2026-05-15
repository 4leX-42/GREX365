using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;

namespace Grex365.App.ViewModels;

public sealed partial class ConnectViewModel : ObservableObject
{
    private readonly IGraphConnection _graph;
    private readonly IExchangeConnection _exchange;
    private readonly IConnectionStateMonitor _monitor;
    private readonly ICertConfigStore _certStore;
    private readonly IUiLogSink _log;
    private CancellationTokenSource? _cts;

    [ObservableProperty]
    private bool _graphConnected;

    [ObservableProperty]
    private bool _exchangeConnected;

    [ObservableProperty]
    private string _statusMessage = "Sin conectar.";

    [ObservableProperty]
    private bool _isBusy;

    public ConnectViewModel(
        IGraphConnection graph,
        IExchangeConnection exchange,
        IConnectionStateMonitor monitor,
        ICertConfigStore certStore,
        IUiLogSink log)
    {
        _graph = graph;
        _exchange = exchange;
        _monitor = monitor;
        _certStore = certStore;
        _log = log;

        _monitor.PropertyChanged += OnMonitorChanged;
        SyncFromMonitor();
    }

    [RelayCommand(CanExecute = nameof(CanConnect))]
    private async Task ConnectAsync()
    {
        if (IsBusy)
        {
            return;
        }

        _cts = new CancellationTokenSource();
        IsBusy = true;
        ConnectCommand.NotifyCanExecuteChanged();
        CancelCommand.NotifyCanExecuteChanged();

        try
        {
            var config = await _certStore.LoadAsync(_cts.Token).ConfigureAwait(true);
            if (config is null)
            {
                StatusMessage = "Falta configuración de certificado (exo-app-params.json).";
                _log.Progress.Report(LogEntry.Warn("Connect", StatusMessage));
                return;
            }

            StatusMessage = "Conectando...";
            await _graph.ConnectByCertificateAsync(config, _log.Progress, _cts.Token).ConfigureAwait(true);
            await _exchange.ConnectByCertificateAsync(config, _log.Progress, _cts.Token).ConfigureAwait(true);
            StatusMessage = "Conectado.";
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusMessage = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("Connect", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
            _cts?.Dispose();
            _cts = null;
            ConnectCommand.NotifyCanExecuteChanged();
            CancelCommand.NotifyCanExecuteChanged();
        }
    }

    [RelayCommand(CanExecute = nameof(CanCancel))]
    private void Cancel()
    {
        _cts?.Cancel();
    }

    [RelayCommand]
    private async Task DisconnectAsync()
    {
        try
        {
            await _exchange.DisconnectAsync(_log.Progress).ConfigureAwait(true);
            await _graph.DisconnectAsync().ConfigureAwait(true);
            StatusMessage = "Desconectado.";
        }
        catch (Exception ex)
        {
            _log.Progress.Report(LogEntry.Error("Disconnect", ex.Message, ex));
        }
    }

    private bool CanConnect() => !IsBusy;

    private bool CanCancel() => IsBusy;

    private void OnMonitorChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        SyncFromMonitor();
    }

    private void SyncFromMonitor()
    {
        var state = _monitor.Current;
        GraphConnected = state.GraphConnected;
        ExchangeConnected = state.ExchangeConnected;
    }
}
