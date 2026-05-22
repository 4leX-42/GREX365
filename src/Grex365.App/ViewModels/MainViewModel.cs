using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Grex365.Core.Plugins;
using Microsoft.Extensions.DependencyInjection;

namespace Grex365.App.ViewModels;

public sealed partial class NavigationItem : ObservableObject
{
    public NavigationItem(string title, string glyph, Type viewModelType, string category = "Otros")
    {
        Title = title;
        Glyph = glyph;
        ViewModelType = viewModelType;
        Category = category;
    }

    public string Title { get; }
    public string Glyph { get; }
    public Type ViewModelType { get; }
    public string Category { get; }
    public bool RequiresGraph { get; set; }
    public bool RequiresExchange { get; set; }

    [ObservableProperty] private bool _isEnabled = true;
    [ObservableProperty] private int _errorBadge;

    public bool HasErrorBadge => ErrorBadge > 0;
    partial void OnErrorBadgeChanged(int value) => OnPropertyChanged(nameof(HasErrorBadge));
}

public sealed partial class MainViewModel : ObservableObject
{
    private readonly IServiceProvider _services;
    private readonly IUiLogSink _uiLog;
    private readonly IConnectionStateMonitor _monitor;
    private readonly IPreferencesStore _prefs;
    private readonly IGraphConnection _graph;
    private readonly IExchangeConnection _exchange;
    private readonly IAuditFindingsStore _auditStore;
    private readonly IUserDetailsHost _userDetailsHost;

    [ObservableProperty] private bool _userDrawerOpen;
    public UserDetailsViewModel UserDetailsVm { get; }

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
    [ObservableProperty] private bool _logPanelVisible;

    public ICollectionView LogView { get; }

    public MainViewModel(
        IUiLogSink uiLog,
        IServiceProvider services,
        IConnectionStateMonitor monitor,
        IPreferencesStore prefs,
        IGraphConnection graph,
        IExchangeConnection exchange,
        IAuditFindingsStore auditStore,
        IUserDetailsHost userDetailsHost,
        UserDetailsViewModel userDetailsVm,
        PluginLoadReport pluginReport)
    {
        _uiLog = uiLog;
        LogEntries = uiLog.Entries;
        _services = services;
        _monitor = monitor;
        _prefs = prefs;
        _graph = graph;
        _exchange = exchange;
        _auditStore = auditStore;
        _userDetailsHost = userDetailsHost;
        UserDetailsVm = userDetailsVm;

        _monitor.PropertyChanged += OnMonitorChanged;
        _auditStore.PropertyChanged += OnAuditStoreChanged;
        _userDetailsHost.PropertyChanged += OnUserDetailsHostChanged;
        SyncFromMonitor();

        try
        {
            var initialPrefs = _prefs.LoadAsync().GetAwaiter().GetResult();
            _logPanelVisible = initialPrefs.LogPanelVisible;
        }
        catch
        {
            _logPanelVisible = false;
        }

        LogView = CollectionViewSource.GetDefaultView(uiLog.Entries);
        LogView.Filter = FilterLogEntry;

        NavigationItems = new ObservableCollection<NavigationItem>
        {
            new("Dashboard",     "", typeof(DashboardViewModel),     "Tenant"),
            new("Conexión",      "", typeof(ConnectViewModel),       "Tenant"),
            new("Salud tenant",  "", typeof(TenantHealthViewModel),  "Tenant"),
            new("Usuarios",      "", typeof(UsersViewModel),         "Identidad"),
            new("Grupos",        "", typeof(GroupsViewModel),        "Identidad"),
            new("Onboarding",    "", typeof(OnboardingViewModel),   "Identidad"),
            new("Offboarding",   "", typeof(OffboardingViewModel),  "Identidad"),
            new("Buzones",       "", typeof(SharedMailboxViewModel), "Mail"),
            new("Reglas de buzón","", typeof(MailboxRulesViewModel),  "Mail"),
            new("Flujo de correo","", typeof(MailFlowRulesViewModel), "Mail"),
            new("Auditoría",     "", typeof(AuditViewModel),         "Seguridad"),
            new("Registro de auditoría","", typeof(AuditLogViewModel),      "Seguridad"),
            new("Asistente cert","", typeof(CertWizardViewModel),   "Herramientas"),
            new("Comprobación DNS","", typeof(DomainCheckViewModel),  "Herramientas"),
        };

        foreach (var module in pluginReport.AllModules)
        {
            NavigationItems.Add(new NavigationItem(module.Title, module.Glyph, module.ViewModelType, "Plugins"));
        }

        NavigationItemsView = CollectionViewSource.GetDefaultView(NavigationItems);
        NavigationItemsView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(NavigationItem.Category)));

        ApplyConnectionRequirements();
        UpdateNavEnabledStates();

        SelectedNavigation = LoadLastNavigation() ?? NavigationItems[0];
    }

    private static readonly HashSet<string> RequiresGraphTitles = new(StringComparer.OrdinalIgnoreCase)
    {
        "Salud tenant",
        "Usuarios",
        "Grupos",
        "Auditoría",
        "Onboarding",
        "Offboarding",
    };

    private static readonly HashSet<string> RequiresExchangeTitles = new(StringComparer.OrdinalIgnoreCase)
    {
        "Buzones",
        "Reglas de buzón",
        "Flujo de correo",
    };

    // Old → new title rename map. Lets users that had old nav persisted keep their last selection.
    private static readonly Dictionary<string, string> RenamedNavTitles = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Conexion"] = "Conexión",
        ["Auditoria"] = "Auditoría",
        ["Reglas buzon"] = "Reglas de buzón",
    };

    private void ApplyConnectionRequirements()
    {
        foreach (var item in NavigationItems)
        {
            if (RequiresGraphTitles.Contains(item.Title))
            {
                item.RequiresGraph = true;
            }
            if (RequiresExchangeTitles.Contains(item.Title))
            {
                item.RequiresExchange = true;
            }
        }
    }

    private void UpdateNavEnabledStates()
    {
        if (NavigationItems is null)
        {
            return;
        }
        foreach (var item in NavigationItems)
        {
            item.IsEnabled =
                (!item.RequiresGraph || GraphConnected) &&
                (!item.RequiresExchange || ExchangeConnected);
        }
    }

    public ObservableCollection<NavigationItem> NavigationItems { get; }
    public ICollectionView NavigationItemsView { get; private set; } = default!;

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
            var target = RenamedNavTitles.TryGetValue(prefs.LastSelectedNavigation, out var renamed)
                ? renamed
                : prefs.LastSelectedNavigation;
            return NavigationItems.FirstOrDefault(i =>
                string.Equals(i.Title, target, StringComparison.OrdinalIgnoreCase));
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
        UpdateNavEnabledStates();
    }

    private void OnUserDetailsHostChanged(object? sender, PropertyChangedEventArgs e)
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is not null && !dispatcher.CheckAccess())
        {
            dispatcher.Invoke(() => UserDrawerOpen = _userDetailsHost.IsOpen);
        }
        else
        {
            UserDrawerOpen = _userDetailsHost.IsOpen;
        }
    }

    private void OnAuditStoreChanged(object? sender, PropertyChangedEventArgs e)
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is not null && !dispatcher.CheckAccess())
        {
            dispatcher.Invoke(SyncAuditBadge);
        }
        else
        {
            SyncAuditBadge();
        }
    }

    private void SyncAuditBadge()
    {
        var item = NavigationItems.FirstOrDefault(i =>
            string.Equals(i.Title, "Auditoría", StringComparison.OrdinalIgnoreCase));
        if (item is not null)
        {
            item.ErrorBadge = _auditStore.ErrorCount;
        }
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
    private async Task ToggleThemeAsync()
    {
        try
        {
            var prefs = await _prefs.LoadAsync().ConfigureAwait(true);
            var current = string.IsNullOrWhiteSpace(prefs.Theme) ? "Dark" : prefs.Theme;
            var next = string.Equals(current, "Dark", StringComparison.OrdinalIgnoreCase) ? "Light" : "Dark";
            prefs.Theme = next;
            await _prefs.SaveAsync(prefs).ConfigureAwait(true);
            SettingsViewModel.ApplyThemeFromPreferences(next);
            _uiLog.Progress.Report(LogEntry.Info("Theme", "Tema cambiado a " + next));
        }
        catch (Exception ex)
        {
            _uiLog.Progress.Report(LogEntry.Error("Theme", ex.Message, ex));
        }
    }

    [RelayCommand]
    private void ClearLog() => _uiLog.Clear();

    partial void OnLogPanelVisibleChanged(bool value)
    {
        _ = PersistLogPanelStateAsync(value);
    }

    private async Task PersistLogPanelStateAsync(bool visible)
    {
        try
        {
            var p = await _prefs.LoadAsync().ConfigureAwait(false);
            if (p.LogPanelVisible == visible)
            {
                return;
            }
            p.LogPanelVisible = visible;
            await _prefs.SaveAsync(p).ConfigureAwait(false);
        }
        catch
        {
            // Ignore — preference is non-critical.
        }
    }

    [RelayCommand]
    private void ToggleLogPanel() => LogPanelVisible = !LogPanelVisible;

    [RelayCommand]
    private void CloseUserDrawer()
    {
        if (UserDrawerOpen)
        {
            _userDetailsHost.RequestClose();
        }
    }

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
