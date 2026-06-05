using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.Core.Abstractions;
using Grex365.Core.Audit;
using Microsoft.Extensions.DependencyInjection;

namespace Grex365.App.ViewModels;

public sealed partial class DashboardViewModel : ObservableObject
{
    private readonly IConnectionStateMonitor _monitor;
    private readonly IAuditFindingsStore _auditStore;
    private readonly IServiceProvider _services;
    private readonly IAuditLog? _auditLog;

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

    [ObservableProperty] private int _todayOpsCount;
    [ObservableProperty] private int _todayErrorCount;
    [ObservableProperty] private bool _hasActivity;
    public ObservableCollection<AuditRecord> RecentOps { get; } = new();

    // auditLog es opcional: sin él, la tarjeta de actividad simplemente no aparece.
    public DashboardViewModel(
        IConnectionStateMonitor monitor,
        IAuditFindingsStore auditStore,
        IServiceProvider services,
        IAuditLog? auditLog = null)
    {
        _monitor = monitor;
        _auditStore = auditStore;
        _services = services;
        _auditLog = auditLog;
        _monitor.PropertyChanged += OnMonitorChanged;
        _auditStore.PropertyChanged += OnAuditStoreChanged;
        Sync();
        SyncAudit();
    }

    // Invocado desde el Loaded de la vista: refresca el resumen local de actividad (audit JSONL
    // del mes en curso; si trae poco — p.ej. día 1 — completa con el mes anterior).
    [RelayCommand]
    private async Task RefreshActivityAsync()
    {
        if (_auditLog is null) return;
        try
        {
            var now = DateTimeOffset.Now;
            var records = (await _auditLog.ReadMonthAsync(now.Year, now.Month).ConfigureAwait(true)).ToList();
            if (records.Count < 5)
            {
                var prev = now.AddMonths(-1);
                records.AddRange(await _auditLog.ReadMonthAsync(prev.Year, prev.Month).ConfigureAwait(true));
            }
            var activity = RecentActivityAggregator.Compute(records, now);
            TodayOpsCount = activity.TodayCount;
            TodayErrorCount = activity.TodayErrors;
            RecentOps.Clear();
            foreach (var r in activity.Recent) RecentOps.Add(r);
            HasActivity = RecentOps.Count > 0;
        }
        catch
        {
            // Best-effort: un JSONL corrupto/inaccesible no debe romper el Dashboard.
            HasActivity = false;
        }
    }

    [RelayCommand]
    private void GoTo(string target)
    {
        var main = _services.GetRequiredService<MainViewModel>();
        // Match by stable NavKey first (i18n-safe); fall back to localized Title.
        var matching = main.NavigationItems.FirstOrDefault(i =>
            string.Equals(i.NavKey, target, StringComparison.OrdinalIgnoreCase))
            ?? main.NavigationItems.FirstOrDefault(i =>
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
