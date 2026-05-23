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
    // navKey + categoryKey son identity stable across language switches.
    // title + category derive from L10n lookup en construction (snapshot at boot).
    public NavigationItem(string title, string glyph, Type viewModelType, string category = "Otros", string? navKey = null, string? categoryKey = null)
    {
        Title = title;
        Glyph = glyph;
        ViewModelType = viewModelType;
        Category = category;
        NavKey = navKey ?? title;
        CategoryKey = categoryKey ?? category;
    }

    public string Title { get; }
    public string Glyph { get; }
    public Type ViewModelType { get; }
    public string Category { get; }
    public string NavKey { get; }
    public string CategoryKey { get; }
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
            BuildNav("Nav.Dashboard", "", typeof(DashboardViewModel), "NavCategory.Tenant"),
            BuildNav("Nav.Connection", "", typeof(ConnectViewModel), "NavCategory.Tenant"),
            BuildNav("Nav.Licenses", "", typeof(TenantHealthViewModel), "NavCategory.Tenant"),
            BuildNav("Nav.Users", "", typeof(UsersViewModel), "NavCategory.Identity"),
            BuildNav("Nav.Groups", "", typeof(GroupsViewModel), "NavCategory.Identity"),
            BuildNav("Nav.Onboarding", "", typeof(OnboardingViewModel), "NavCategory.Identity"),
            BuildNav("Nav.Offboarding", "", typeof(OffboardingViewModel), "NavCategory.Identity"),
            BuildNav("Nav.SharedMailbox", "", typeof(SharedMailboxViewModel), "NavCategory.Mail"),
            BuildNav("Nav.MailboxRules", "", typeof(MailboxRulesViewModel), "NavCategory.Mail"),
            BuildNav("Nav.MailFlow", "", typeof(MailFlowRulesViewModel), "NavCategory.Mail"),
            BuildNav("Nav.Audit", "", typeof(AuditViewModel), "NavCategory.Security"),
            BuildNav("Nav.AuditLog", "", typeof(AuditLogViewModel), "NavCategory.Security"),
            BuildNav("Nav.PsConsole", "", typeof(PsConsoleViewModel), "NavCategory.Tools"),
            BuildNav("Nav.CertWizard", "", typeof(CertWizardViewModel), "NavCategory.Tools"),
            BuildNav("Nav.DnsCheck", "", typeof(DomainCheckViewModel), "NavCategory.Tools"),
        };

        foreach (var module in pluginReport.AllModules)
        {
            NavigationItems.Add(new NavigationItem(
                title: module.Title,
                glyph: module.Glyph,
                viewModelType: module.ViewModelType,
                category: L10n.Get("NavCategory.Plugins"),
                navKey: "Plugin." + module.Title,
                categoryKey: "NavCategory.Plugins"));
        }

        NavigationItemsView = CollectionViewSource.GetDefaultView(NavigationItems);
        NavigationItemsView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(NavigationItem.Category)));

        ApplyConnectionRequirements();
        UpdateNavEnabledStates();

        SelectedNavigation = LoadLastNavigation() ?? NavigationItems[0];
    }

    private static readonly HashSet<string> RequiresGraphKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "Nav.Licenses",
        "Nav.Users",
        "Nav.Groups",
        "Nav.Audit",
        "Nav.Onboarding",
        "Nav.Offboarding",
    };

    private static readonly HashSet<string> RequiresExchangeKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "Nav.SharedMailbox",
        "Nav.MailboxRules",
        "Nav.MailFlow",
        "Nav.Audit",
    };

    private static NavigationItem BuildNav(string navKey, string glyph, Type viewModelType, string categoryKey)
    {
        return new NavigationItem(
            title: L10n.Get(navKey),
            glyph: glyph,
            viewModelType: viewModelType,
            category: L10n.Get(categoryKey),
            navKey: navKey,
            categoryKey: categoryKey);
    }

    private void ApplyConnectionRequirements()
    {
        foreach (var item in NavigationItems)
        {
            if (RequiresGraphKeys.Contains(item.NavKey))
            {
                item.RequiresGraph = true;
            }
            if (RequiresExchangeKeys.Contains(item.NavKey))
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
            item.IsEnabled = NavRequirements.IsEnabled(item, GraphConnected, ExchangeConnected);
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
            var saved = prefs.LastSelectedNavigation;
            if (string.IsNullOrWhiteSpace(saved)) return null;

            // Modern: saved value is the NavKey (language-stable).
            var byKey = NavigationItems.FirstOrDefault(i =>
                string.Equals(i.NavKey, saved, StringComparison.OrdinalIgnoreCase));
            if (byKey is not null) return byKey;

            // Legacy: saved value is a Title (possibly pre-rename). Migrate then match by current Title.
            var target = NavTitleMigrator.Resolve(saved);
            if (target is null) return null;
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
        _ = PersistNavAsync(value.NavKey);
    }

    private async Task PersistNavAsync(string navKey)
    {
        try
        {
            var p = await _prefs.LoadAsync().ConfigureAwait(false);
            if (string.Equals(p.LastSelectedNavigation, navKey, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
            p.LastSelectedNavigation = navKey;
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
    private void OpenAbout()
    {
        var window = _services.GetRequiredService<AboutWindow>();
        window.Owner = Application.Current?.MainWindow;
        window.ShowDialog();
    }

    [RelayCommand]
    private async Task ExportLogAsync()
    {
        try
        {
            var stamp = DateTimeOffset.Now.ToString("yyyyMMdd-HHmmss");
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Exportar log",
                FileName = $"grex365-log-{stamp}.txt",
                DefaultExt = ".txt",
                Filter = "Texto (*.txt)|*.txt|CSV (*.csv)|*.csv|Todos|*.*",
                AddExtension = true
            };
            if (dlg.ShowDialog() != true)
            {
                return;
            }

            var path = dlg.FileName;
            var entries = new List<LogEntry>();
            foreach (var obj in LogView)
            {
                if (obj is LogEntry e)
                {
                    entries.Add(e);
                }
            }

            var isCsv = path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase);
            using var writer = new System.IO.StreamWriter(path, append: false, System.Text.Encoding.UTF8);
            if (isCsv)
            {
                await writer.WriteLineAsync("Timestamp,Severity,Source,Message").ConfigureAwait(false);
                foreach (var e in entries)
                {
                    var ts = e.Timestamp.ToString("yyyy-MM-dd HH:mm:ss");
                    await writer.WriteLineAsync(
                        $"{ts},{e.Severity},{CsvEscape(e.Source)},{CsvEscape(e.Message)}").ConfigureAwait(false);
                }
            }
            else
            {
                foreach (var e in entries)
                {
                    var ts = e.Timestamp.ToString("yyyy-MM-dd HH:mm:ss");
                    await writer.WriteLineAsync($"{ts} [{e.Severity,-5}] {e.Source}: {e.Message}").ConfigureAwait(false);
                }
            }
            _uiLog.Progress.Report(LogEntry.Ok("Log", $"Log exportado ({entries.Count} entradas) → {path}"));
        }
        catch (Exception ex)
        {
            _uiLog.Progress.Report(LogEntry.Error("Log", "Error exportando: " + ex.Message, ex));
        }
    }

    private static string CsvEscape(string? value) => Grex365.Core.Csv.CsvEscaper.Escape(value);

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
