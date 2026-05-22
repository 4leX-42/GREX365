using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.Core.Abstractions;
using Microsoft.Extensions.DependencyInjection;

namespace Grex365.App.ViewModels;

public sealed partial class DashboardViewModel : ObservableObject
{
    private readonly IConnectionStateMonitor _monitor;
    private readonly IAuditFindingsStore _auditStore;
    private readonly IServiceProvider _services;

    [ObservableProperty] private bool _graphConnected;
    [ObservableProperty] private bool _exchangeConnected;
    [ObservableProperty] private string? _tenantId;
    [ObservableProperty] private string? _tenantDomain;
    [ObservableProperty] private string? _account;
    [ObservableProperty] private string? _lastAuditName;
    [ObservableProperty] private DateTimeOffset? _lastAuditAt;
    [ObservableProperty] private int _lastAuditErrors;
    [ObservableProperty] private int _lastAuditWarnings;
    [ObservableProperty] private int _lastAuditInfo;
    [ObservableProperty] private bool _hasLastAudit;

    public DashboardViewModel(
        IConnectionStateMonitor monitor,
        IAuditFindingsStore auditStore,
        IServiceProvider services)
    {
        _monitor = monitor;
        _auditStore = auditStore;
        _services = services;
        _monitor.PropertyChanged += OnMonitorChanged;
        _auditStore.PropertyChanged += OnAuditStoreChanged;
        Sync();
        SyncAudit();
    }

    [RelayCommand]
    private void GoTo(string target)
    {
        var main = _services.GetRequiredService<MainViewModel>();
        // resolve the same instance App uses
        var matching = main.NavigationItems.FirstOrDefault(i =>
            string.Equals(i.Title, target, StringComparison.OrdinalIgnoreCase));
        if (matching is not null)
        {
            main.SelectedNavigation = matching;
        }
    }

    private void OnMonitorChanged(object? sender, PropertyChangedEventArgs e)
    {
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher is not null && !dispatcher.CheckAccess())
        {
            dispatcher.Invoke(Sync);
        }
        else
        {
            Sync();
        }
    }

    private void Sync()
    {
        var s = _monitor.Current;
        GraphConnected = s.GraphConnected;
        ExchangeConnected = s.ExchangeConnected;
        TenantId = s.TenantId;
        TenantDomain = s.TenantDomain;
        Account = s.Account;
    }

    private void OnAuditStoreChanged(object? sender, PropertyChangedEventArgs e)
    {
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher is not null && !dispatcher.CheckAccess())
        {
            dispatcher.Invoke(SyncAudit);
        }
        else
        {
            SyncAudit();
        }
    }

    private void SyncAudit()
    {
        LastAuditName = _auditStore.LastAuditName;
        LastAuditAt = _auditStore.LastRunAt;
        LastAuditErrors = _auditStore.ErrorCount;
        LastAuditWarnings = _auditStore.WarnCount;
        LastAuditInfo = _auditStore.InfoCount;
        HasLastAudit = _auditStore.LastRunAt.HasValue;
    }
}
