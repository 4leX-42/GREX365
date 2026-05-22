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
    private readonly IRbacGuard _rbac;
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

    [ObservableProperty] private string? _deviceCodeUserCode;
    [ObservableProperty] private string? _deviceCodeVerificationUri;
    [ObservableProperty] private string? _deviceCodeMessage;
    [ObservableProperty] private bool _deviceCodePromptVisible;

    [ObservableProperty] private string _exoModuleStatus = "(no comprobado)";
    [ObservableProperty] private bool _exoModuleAvailable;

    [ObservableProperty] private string? _certAppId;
    [ObservableProperty] private string? _certThumbprint;
    [ObservableProperty] private string? _certExpiry;
    [ObservableProperty] private string? _certStatusMessage;
    [ObservableProperty] private bool _certIsValid;

    public ConnectViewModel(
        IGraphConnection graph,
        IExchangeConnection exchange,
        IConnectionStateMonitor monitor,
        ICertConfigStore certStore,
        ICertValidator certValidator,
        ITenantLock tenantLock,
        IUiLogSink log,
        IRbacGuard rbac)
    {
        _graph = graph;
        _exchange = exchange;
        _monitor = monitor;
        _certStore = certStore;
        _certValidator = certValidator;
        _tenantLock = tenantLock;
        _log = log;
        _rbac = rbac;

        _monitor.PropertyChanged += OnMonitorChanged;
        _ = LoadCertInfoAsync();
        SyncFromMonitor();
    }

    private async Task LoadCertInfoAsync()
    {
        try
        {
            var config = await _certStore.LoadAsync().ConfigureAwait(true);
            if (config is null)
            {
                CertStatusMessage = "Sin config (exo-app-params.json no encontrado).";
                CertIsValid = false;
                return;
            }
            CertAppId = config.AppId;
            CertThumbprint = config.CertThumbprint;
            var v = _certValidator.Validate(config);
            CertIsValid = v.IsValid;
            CertStatusMessage = v.Message;
            CertExpiry = v.NotAfter?.ToString("yyyy-MM-dd");
        }
        catch (Exception ex)
        {
            CertStatusMessage = "Error cargando cert: " + ex.Message;
            CertIsValid = false;
        }
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

    [RelayCommand(CanExecute = nameof(CanConnect))]
    private async Task ConnectDeviceCodeAsync()
    {
        if (IsBusy)
        {
            return;
        }

        _cts = new CancellationTokenSource();
        IsBusy = true;
        ConnectCommand.NotifyCanExecuteChanged();
        ConnectDeviceCodeCommand.NotifyCanExecuteChanged();
        CancelCommand.NotifyCanExecuteChanged();
        DeviceCodePromptVisible = false;

        try
        {
            var config = await _certStore.LoadAsync(_cts.Token).ConfigureAwait(true);
            var tenantHint = config?.TenantId;

            StatusMessage = "Esperando codigo de dispositivo...";
            await _graph.ConnectByDeviceCodeAsync(
                tenantHint,
                (prompt, ct) =>
                {
                    var dispatcher = System.Windows.Application.Current?.Dispatcher;
                    void apply()
                    {
                        DeviceCodeUserCode = prompt.UserCode;
                        DeviceCodeVerificationUri = prompt.VerificationUri;
                        DeviceCodeMessage = prompt.Message;
                        DeviceCodePromptVisible = true;
                        StatusMessage = $"Pega el codigo {prompt.UserCode} en {prompt.VerificationUri}";
                    }
                    if (dispatcher is not null && !dispatcher.CheckAccess())
                    {
                        dispatcher.Invoke(apply);
                    }
                    else
                    {
                        apply();
                    }
                    return Task.CompletedTask;
                },
                _log.Progress,
                _cts.Token).ConfigureAwait(true);

            if (!string.IsNullOrWhiteSpace(_graph.TenantId))
            {
                try
                {
                    await _tenantLock.EnforceAsync(_graph.TenantId, _cts.Token).ConfigureAwait(true);
                }
                catch (TenantLockViolationException violation)
                {
                    _log.Progress.Report(LogEntry.Error("TenantLock", violation.Message, violation));
                    await _graph.DisconnectAsync(_cts.Token).ConfigureAwait(true);
                    StatusMessage = "Tenant lock: " + violation.Message;
                    return;
                }
            }

            DeviceCodePromptVisible = false;
            StatusMessage = $"Conectado a Graph como {_graph.Account ?? "?"}.";
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
            DeviceCodePromptVisible = false;
            _cts?.Dispose();
            _cts = null;
            ConnectCommand.NotifyCanExecuteChanged();
            ConnectDeviceCodeCommand.NotifyCanExecuteChanged();
            CancelCommand.NotifyCanExecuteChanged();
        }
    }

    [RelayCommand]
    private async Task ProbeExoModuleAsync()
    {
        try
        {
            var status = await _exchange.ProbeModuleAsync(_log.Progress).ConfigureAwait(true);
            ExoModuleAvailable = status.Installed;
            ExoModuleStatus = status.Installed
                ? $"OK · ExchangeOnlineManagement {status.Version}"
                : $"NO instalado ({status.Detail ?? "?"})";
        }
        catch (Exception ex)
        {
            ExoModuleStatus = "Probe error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("EXO", ex.Message, ex));
        }
    }

    [RelayCommand]
    private async Task InstallExoModuleAsync()
    {
        if (IsBusy)
        {
            return;
        }
        IsBusy = true;
        ExoModuleStatus = "Instalando ExchangeOnlineManagement (PSGallery)...";
        try
        {
            var status = await _exchange.InstallModuleAsync(_log.Progress).ConfigureAwait(true);
            ExoModuleAvailable = status.Installed;
            ExoModuleStatus = status.Installed
                ? $"Instalado · {status.Version}"
                : $"Instalacion fallida: {status.Detail ?? "?"}";
        }
        catch (Exception ex)
        {
            ExoModuleStatus = "Error: " + ex.Message;
            _log.Progress.Report(LogEntry.Error("EXO", ex.Message, ex));
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task DisconnectAsync()
    {
        try
        {
            await _exchange.DisconnectAsync(_log.Progress).ConfigureAwait(true);
            await _graph.DisconnectAsync().ConfigureAwait(true);
            _rbac.Invalidate();
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
