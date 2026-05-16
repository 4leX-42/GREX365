using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using Grex365.Core.Abstractions;

namespace Grex365.App.ViewModels;

public sealed partial class DashboardViewModel : ObservableObject
{
    private readonly IConnectionStateMonitor _monitor;

    [ObservableProperty] private bool _graphConnected;
    [ObservableProperty] private bool _exchangeConnected;
    [ObservableProperty] private string? _tenantId;
    [ObservableProperty] private string? _tenantDomain;
    [ObservableProperty] private string? _account;

    public DashboardViewModel(IConnectionStateMonitor monitor)
    {
        _monitor = monitor;
        _monitor.PropertyChanged += OnMonitorChanged;
        Sync();
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
}
