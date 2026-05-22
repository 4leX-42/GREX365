using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Extensions.DependencyInjection;

namespace Grex365.App.ViewModels;

public sealed partial class TenantHealthViewModel : ObservableObject
{
    private readonly ITenantHealthService _service;
    private readonly IUiLogSink _log;
    private readonly IConnectionStateMonitor _monitor;
    private readonly IServiceProvider _services;
    private CancellationTokenSource? _cts;
    private bool _autoLoadAttempted;

    [ObservableProperty] private TenantHealth? _health;
    [ObservableProperty] private string _statusMessage = "Conecta Graph para cargar licencias.";
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private int _totalConsumed;
    [ObservableProperty] private int _totalEnabled;
    [ObservableProperty] private double _overallPercent;
    [ObservableProperty] private string _licenseFilter = string.Empty;

    public ObservableCollection<LicenseCard> Licenses { get; } = new();
    public ICollectionView LicensesView { get; }

    public TenantHealthViewModel(
        ITenantHealthService service,
        IUiLogSink log,
        IConnectionStateMonitor monitor,
        IServiceProvider services)
    {
        _service = service;
        _log = log;
        _monitor = monitor;
        _services = services;

        LicensesView = CollectionViewSource.GetDefaultView(Licenses);
        LicensesView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(LicenseCard.CategoryLabel)));
        LicensesView.SortDescriptions.Add(new SortDescription(nameof(LicenseCard.Priority), ListSortDirection.Ascending));
        LicensesView.Filter = LicenseFilterPredicate;

        _monitor.PropertyChanged += OnMonitorChanged;
        TryAutoLoad();
    }

    private bool LicenseFilterPredicate(object obj)
    {
        if (obj is not LicenseCard l) return false;
        var q = (LicenseFilter ?? string.Empty).Trim();
        if (q.Length == 0) return true;
        return (l.FriendlyName?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
            || (l.SkuPartNumber?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
            || (l.CategoryLabel?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false);
    }

    partial void OnLicenseFilterChanged(string value) => LicensesView.Refresh();

    private void OnMonitorChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (!string.Equals(e.PropertyName, nameof(IConnectionStateMonitor.Current), StringComparison.Ordinal))
        {
            return;
        }
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher is null) return;
        dispatcher.InvokeAsync(() => TryAutoLoad());
    }

    private void TryAutoLoad()
    {
        if (_autoLoadAttempted || IsBusy) return;
        if (!_monitor.Current.GraphConnected) return;
        _autoLoadAttempted = true;
        if (RefreshCommand.CanExecute(null))
        {
            RefreshCommand.Execute(null);
        }
    }

    [RelayCommand(CanExecute = nameof(CanRefresh))]
    private async Task RefreshAsync()
    {
        _cts = new CancellationTokenSource();
        IsBusy = true;
        RefreshCommand.NotifyCanExecuteChanged();
        CancelCommand.NotifyCanExecuteChanged();
        StatusMessage = "Cargando licencias del tenant...";

        try
        {
            var h = await _service.GetAsync(_log.Progress, _cts.Token).ConfigureAwait(true);
            Health = h;
            Licenses.Clear();
            foreach (var l in h.Licenses)
            {
                Licenses.Add(LicenseCard.From(l));
            }
            TotalConsumed = h.Licenses.Sum(l => l.Consumed);
            TotalEnabled = h.Licenses.Sum(l => l.Enabled);
            OverallPercent = TotalEnabled > 0 ? (TotalConsumed * 100.0 / TotalEnabled) : 0;
            StatusMessage = $"{h.TotalUsers} usuarios · {h.TotalGroups} grupos · {h.Licenses.Count} SKUs · {TotalConsumed}/{TotalEnabled} asientos consumidos";
            LicensesView.Refresh();
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

    [RelayCommand]
    private void GoToUsers()
    {
        var main = _services.GetRequiredService<MainViewModel>();
        var target = main.NavigationItems.FirstOrDefault(i =>
            string.Equals(i.Title, "Usuarios", StringComparison.OrdinalIgnoreCase));
        if (target is not null)
        {
            main.SelectedNavigation = target;
        }
    }

    [RelayCommand]
    private void ClearFilter() => LicenseFilter = string.Empty;

    private bool CanRefresh() => !IsBusy;
    private bool CanCancel() => IsBusy;
}
