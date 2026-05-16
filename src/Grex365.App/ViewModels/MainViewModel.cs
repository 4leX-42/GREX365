using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Microsoft.Extensions.DependencyInjection;

namespace Grex365.App.ViewModels;

public sealed class NavigationItem
{
    public NavigationItem(string title, string glyph, Type viewModelType)
    {
        Title = title;
        Glyph = glyph;
        ViewModelType = viewModelType;
    }

    public string Title { get; }
    public string Glyph { get; }
    public Type ViewModelType { get; }
}

public sealed partial class MainViewModel : ObservableObject
{
    private readonly IServiceProvider _services;
    private readonly IUiLogSink _uiLog;
    private readonly IConnectionStateMonitor _monitor;
    private readonly IPreferencesStore _prefs;
    private readonly IGraphConnection _graph;
    private readonly IExchangeConnection _exchange;

    [ObservableProperty] private NavigationItem? _selectedNavigation;
    [ObservableProperty] private ObservableObject? _currentPage;

    [ObservableProperty] private bool _graphConnected;
    [ObservableProperty] private bool _exchangeConnected;
    [ObservableProperty] private string? _tenantId;
    [ObservableProperty] private string? _tenantDomain;
    [ObservableProperty] private string? _account;

    [ObservableProperty] private bool _showInfo = true;
    [ObservableProperty] private bool _showOk = true;
    [ObservableProperty] private bool _showWarn = true;
    [ObservableProperty] private bool _showError = true;
    [ObservableProperty] private bool _showDebug;

    public ICollectionView LogView { get; }

    public MainViewModel(
        IUiLogSink uiLog,
        IServiceProvider services,
        IConnectionStateMonitor monitor,
        IPreferencesStore prefs,
        IGraphConnection graph,
        IExchangeConnection exchange)
    {
        _uiLog = uiLog;
        LogEntries = uiLog.Entries;
        _services = services;
        _monitor = monitor;
        _prefs = prefs;
        _graph = graph;
        _exchange = exchange;

        _monitor.PropertyChanged += OnMonitorChanged;
        SyncFromMonitor();

        LogView = CollectionViewSource.GetDefaultView(uiLog.Entries);
        LogView.Filter = FilterLogEntry;

        NavigationItems = new ObservableCollection<NavigationItem>
        {
            new("Dashboard",      "", typeof(DashboardViewModel)),
            new("Conexion",       "", typeof(ConnectViewModel)),
            new("Salud tenant",   "", typeof(TenantHealthViewModel)),
            new("Usuarios",       "", typeof(UsersViewModel)),
            new("Grupos",         "", typeof(GroupsViewModel)),
            new("Buzones",        "", typeof(SharedMailboxViewModel)),
            new("Auditoria",      "", typeof(AuditViewModel)),
            new("Offboarding",    "", typeof(OffboardingViewModel)),
            new("Cert Wizard",    "", typeof(CertWizardViewModel)),
            new("DNS check",      "", typeof(DomainCheckViewModel)),
        };

        SelectedNavigation = LoadLastNavigation() ?? NavigationItems[0];
    }

    public ObservableCollection<NavigationItem> NavigationItems { get; }

    public ObservableCollection<LogEntry> LogEntries { get; }

    private NavigationItem? LoadLastNavigation()
    {
        try
        {
            var prefs = _prefs.LoadAsync().GetAwaiter().GetResult();
            if (string.IsNullOrWhiteSpace(prefs.LastSelectedNavigation))
            {
                return null;
            }
            return NavigationItems.FirstOrDefault(i =>
                string.Equals(i.Title, prefs.LastSelectedNavigation, StringComparison.OrdinalIgnoreCase));
        }
        catch
        {
            return null;
        }
    }

    partial void OnSelectedNavigationChanged(NavigationItem? value)
    {
        if (value is null)
        {
            CurrentPage = null;
            return;
        }
        CurrentPage = (ObservableObject)_services.GetRequiredService(value.ViewModelType);
        _ = PersistNavAsync(value.Title);
    }

    private async Task PersistNavAsync(string title)
    {
        try
        {
            var p = await _prefs.LoadAsync().ConfigureAwait(false);
            if (string.Equals(p.LastSelectedNavigation, title, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
            p.LastSelectedNavigation = title;
            await _prefs.SaveAsync(p).ConfigureAwait(false);
        }
        catch
        {
            // ignore persistence failures
        }
    }

    private void OnMonitorChanged(object? sender, PropertyChangedEventArgs e)
    {
        var dispatcher = Application.Current?.Dispatcher;
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
        var s = _monitor.Current;
        GraphConnected = s.GraphConnected;
        ExchangeConnected = s.ExchangeConnected;
        TenantId = s.TenantId;
        TenantDomain = s.TenantDomain;
        Account = s.Account;
    }

    [RelayCommand]
    private void OpenSettings()
    {
        var window = _services.GetRequiredService<SettingsWindow>();
        window.Owner = Application.Current?.MainWindow;
        window.ShowDialog();
    }

    [RelayCommand]
    private async Task DisconnectAllAsync()
    {
        try
        {
            await _exchange.DisconnectAsync(_uiLog.Progress).ConfigureAwait(true);
            await _graph.DisconnectAsync().ConfigureAwait(true);
            _uiLog.Progress.Report(LogEntry.Info("Connect", "Desconectado de Graph y Exchange Online."));
        }
        catch (Exception ex)
        {
            _uiLog.Progress.Report(LogEntry.Error("Connect", ex.Message, ex));
        }
    }

    [RelayCommand]
    private void ClearLog() => _uiLog.Clear();

    partial void OnShowInfoChanged(bool value) => LogView.Refresh();
    partial void OnShowOkChanged(bool value) => LogView.Refresh();
    partial void OnShowWarnChanged(bool value) => LogView.Refresh();
    partial void OnShowErrorChanged(bool value) => LogView.Refresh();
    partial void OnShowDebugChanged(bool value) => LogView.Refresh();

    private bool FilterLogEntry(object obj)
    {
        if (obj is not LogEntry e)
        {
            return false;
        }
        return e.Severity switch
        {
            LogSeverity.Info => ShowInfo,
            LogSeverity.Ok => ShowOk,
            LogSeverity.Warning => ShowWarn,
            LogSeverity.Error => ShowError,
            LogSeverity.Debug => ShowDebug,
            _ => true
        };
    }
}
