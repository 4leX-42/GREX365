using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.Core.Abstractions;
using Microsoft.Extensions.DependencyInjection;

namespace Grex365.App.ViewModels;

public sealed partial class DashboardViewModel : ObservableObject
{
    private readonly IConnectionStateMonitor _monitor;
    private readonly IServiceProvider _services;

    [ObservableProperty] private bool _graphConnected;
    [ObservableProperty] private bool _exchangeConnected;
    [ObservableProperty] private string? _tenantId;
    [ObservableProperty] private string? _tenantDomain;
    [ObservableProperty] private string? _account;

    public DashboardViewModel(IConnectionStateMonitor monitor, IServiceProvider services)
    {
        _monitor = monitor;
        _services = services;
        _monitor.PropertyChanged += OnMonitorChanged;
        Sync();
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
}
