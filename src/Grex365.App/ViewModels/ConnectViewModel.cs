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
    private readonly ICertValidator _certValidator;
    private readonly ITenantLock _tenantLock;
    private readonly IUiLogSink _log;
    private CancellationTokenSource? _cts;

    [ObservableProperty]
    private bool _graphConnected;

    [ObservableProperty]
    private bool _exchangeConnected;

    [ObservableProperty]
    private string? _tenantId;

    [ObservableProperty]
    private string? _tenantDomain;

    [ObservableProperty]
    private string? _account;

    [ObservableProperty]
    private string _statusMessage = "Sin conectar.";

    [ObservableProperty]
    private bool _isBusy;

    public ConnectViewModel(
        IGraphConnection graph,
        IExchangeConnection exchange,
        IConnectionStateMonitor monitor,
        ICertConfigStore certStore,
        ICertValidator certValidator,
        ITenantLock tenantLock,
        IUiLogSink log)
    {
        _graph = graph;
        _exchange = exchange;
        _monitor = monitor;
        _certStore = certStore;
        _certValidator = certValidator;
        _tenantLock = tenantLock;
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

            var validation = _certValidator.Validate(config);
            if (!validation.IsValid)
            {
                StatusMessage = "Cert inválido: " + validation.Message;
                _log.Progress.Report(LogEntry.Error("Connect", StatusMessage));
                return;
            }
            _log.Progress.Report(LogEntry.Info("Connect", validation.Message));

            StatusMessage = "Conectando...";
            await _graph.ConnectByCertificateAsync(config, _log.Progress, _cts.Token).ConfigureAwait(true);

            try
            {
                await _tenantLock.EnforceAsync(_graph.TenantId ?? config.TenantId, _cts.Token).ConfigureAwait(true);
            }
            catch (TenantLockViolationException violation)
            {
                _log.Progress.Report(LogEntry.Error("TenantLock", violation.Message, violation));
                await _graph.DisconnectAsync(_cts.Token).ConfigureAwait(true);
                StatusMessage = "Tenant lock: " + violation.Message;
                return;
            }

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
        // Marshal to UI thread; monitor PropertyChanged fires from background task.
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher is not null && !dispatcher.CheckAccess())
        {
            dispatcher.Invoke(SyncFromMonitor);
        }
        else
        {
            SyncFromMonitor();
        }
    }

    private void SyncFromMonitor()
    {
        var state = _monitor.Current;
        GraphConnected = state.GraphConnected;
        ExchangeConnected = state.ExchangeConnected;
        TenantId = state.TenantId;
        TenantDomain = state.TenantDomain;
        Account = state.Account;
    }
}
