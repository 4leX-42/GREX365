using System.Collections.ObjectModel;
using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Grex365.App.Services;
using Grex365.Core.Abstractions;
using Grex365.Core.Models;
using Grex365.Core.Plugins;
using Serilog.Core;

namespace Grex365.App.ViewModels;

public sealed partial class PluginToggleItem : ObservableObject
{
    public string AssemblyFileName { get; }
    public string Status { get; }
    public int ModuleCount { get; }
    public string ModuleCountLabel => $"{ModuleCount} {L10n.Get("Settings.Plugins.ModulesSuffix")}";
    [ObservableProperty] private bool _isEnabled;

    public PluginToggleItem(string fileName, string status, int moduleCount, bool isEnabled)
    {
        AssemblyFileName = fileName;
        Status = status;
        ModuleCount = moduleCount;
        IsEnabled = isEnabled;
    }
}

public sealed partial class SettingsViewModel : ObservableObject
{
    private readonly IPreferencesStore _prefsStore;
    private readonly ICertConfigStore _certStore;
    private readonly ICertValidator _certValidator;
    private readonly IUiLogSink _log;
    private readonly PluginLoadReport _pluginReport;
    private readonly LoggingLevelSwitch _logLevelSwitch;

    [ObservableProperty] private string _connectionMethod = "cert";
    [ObservableProperty] private string? _expectedTenantId;
    [ObservableProperty] private string? _expectedTenantDomain;
    [ObservableProperty] private bool _enforceTenantLock;
    [ObservableProperty] private string _theme = "Dark";
    [ObservableProperty] private string _language = "es";
    [ObservableProperty] private string _languageRestartHint = string.Empty;
    private string _initialLanguage = "es";
    [ObservableProperty] private string _logLevel = "Information";
    [ObservableProperty] private string? _applicationInsightsConnectionString;
    [ObservableProperty] private string? _authorizationGroupId;

    [ObservableProperty] private string _certAppId = string.Empty;
    [ObservableProperty] private string _certTenantId = string.Empty;
    [ObservableProperty] private string _certOrganization = string.Empty;
    [ObservableProperty] private string _certThumbprint = string.Empty;

    [ObservableProperty] private string _certStatusMessage = "—";
    [ObservableProperty] private bool _certIsValid;

    [ObservableProperty] private string _saveStatus = string.Empty;

    public ObservableCollection<PluginToggleItem> Plugins { get; } = new();

    public SettingsViewModel(
        IPreferencesStore prefsStore,
        ICertConfigStore certStore,
        ICertValidator certValidator,
        IUiLogSink log,
        PluginLoadReport pluginReport,
        LoggingLevelSwitch logLevelSwitch)
    {
        _prefsStore = prefsStore;
        _certStore = certStore;
        _certValidator = certValidator;
        _log = log;
        _pluginReport = pluginReport;
        _logLevelSwitch = logLevelSwitch;
    }

    [RelayCommand]
    private async Task LoadAsync()
    {
        var prefs = await _prefsStore.LoadAsync().ConfigureAwait(true);
        ConnectionMethod = prefs.ConnectionMethod ?? "cert";
        ExpectedTenantId = prefs.ExpectedTenantId;
        ExpectedTenantDomain = prefs.ExpectedTenantDomain;
        EnforceTenantLock = prefs.EnforceTenantLock;
        Theme = string.IsNullOrWhiteSpace(prefs.Theme) ? "Dark" : prefs.Theme;
        Language = string.IsNullOrWhiteSpace(prefs.Language) ? "es" : prefs.Language;
        _initialLanguage = Language;
        LanguageRestartHint = string.Empty;
        LogLevel = string.IsNullOrWhiteSpace(prefs.LogLevel) ? "Information" : prefs.LogLevel;
        ApplicationInsightsConnectionString = prefs.ApplicationInsightsConnectionString;
        AuthorizationGroupId = prefs.AuthorizationGroupId;

        var cert = await _certStore.LoadAsync().ConfigureAwait(true);
        if (cert is not null)
        {
            CertAppId = cert.AppId;
            CertTenantId = cert.TenantId;
            CertOrganization = cert.Organization;
            CertThumbprint = cert.CertThumbprint;
        }

        RebuildPluginList(prefs);
        ValidateCert();
    }

    private void RebuildPluginList(UserPreferences prefs)
    {
        Plugins.Clear();
        foreach (var p in _pluginReport.Plugins)
        {
            Plugins.Add(new PluginToggleItem(
                Path.GetFileName(p.AssemblyPath),
                L10n.Get("Settings.Plugins.Status.Loaded"),
                p.Modules.Count,
                isEnabled: true));
        }
        foreach (var f in _pluginReport.Failures)
        {
            Plugins.Add(new PluginToggleItem(
                Path.GetFileName(f.AssemblyPath),
                L10n.Get("Settings.Plugins.Status.ErrorPrefix") + f.Message,
                0,
                isEnabled: true));
        }
        foreach (var d in _pluginReport.Disabled)
        {
            Plugins.Add(new PluginToggleItem(
                d.AssemblyFileName,
                L10n.Get("Settings.Plugins.Status.Disabled"),
                0,
                isEnabled: false));
        }
    }

    [RelayCommand]
    private async Task SaveAsync()
    {
        try
        {
            var prefs = await _prefsStore.LoadAsync().ConfigureAwait(true);
            prefs.ConnectionMethod = ConnectionMethod;
            prefs.ExpectedTenantId = ExpectedTenantId;
            prefs.ExpectedTenantDomain = ExpectedTenantDomain;
            prefs.EnforceTenantLock = EnforceTenantLock;
            prefs.Theme = Theme;
            prefs.Language = string.IsNullOrWhiteSpace(Language) ? "es" : Language.Trim().ToLowerInvariant();
            prefs.LogLevel = LogLevel;
            prefs.ApplicationInsightsConnectionString = string.IsNullOrWhiteSpace(ApplicationInsightsConnectionString)
                ? null
                : ApplicationInsightsConnectionString.Trim();
            prefs.AuthorizationGroupId = string.IsNullOrWhiteSpace(AuthorizationGroupId)
                ? null
                : AuthorizationGroupId.Trim();
            _logLevelSwitch.MinimumLevel = App.ParseLogLevel(LogLevel);
            prefs.DisabledPluginAssemblies = Plugins
                .Where(p => !p.IsEnabled)
                .Select(p => p.AssemblyFileName)
                .ToList();
            await _prefsStore.SaveAsync(prefs).ConfigureAwait(true);

            ApplyTheme(Theme);

            if (!string.IsNullOrWhiteSpace(CertAppId)
                && !string.IsNullOrWhiteSpace(CertTenantId)
                && !string.IsNullOrWhiteSpace(CertOrganization)
                && !string.IsNullOrWhiteSpace(CertThumbprint))
            {
                await _certStore.SaveAsync(new CertConfig(
                    CertAppId, CertTenantId, CertOrganization, CertThumbprint))
                    .ConfigureAwait(true);
            }

            SaveStatus = L10n.Format("Settings.SaveStatus.Prefix", DateTime.Now.ToString("HH:mm:ss"));
            _log.Progress.Report(LogEntry.Ok("Settings", L10n.Get("Settings.SaveStatus.SavedLog")));

            if (!string.Equals(prefs.Language, _initialLanguage, StringComparison.OrdinalIgnoreCase))
            {
                LanguageRestartHint = L10n.Get("Settings.RestartRequired");
            }

            ValidateCert();
        }
        catch (Exception ex)
        {
            SaveStatus = L10n.Format("Settings.SaveStatus.ErrorPrefix", ex.Message);
            _log.Progress.Report(LogEntry.Error("Settings", ex.Message, ex));
        }
    }

    // Set by App at startup. Null fallback returns dark.
    public static ISystemThemeProvider? SystemThemeProvider { get; set; }

    private static void ApplyTheme(string theme)
    {
        Wpf.Ui.Appearance.ApplicationTheme resolved;
        if (string.Equals(theme, "Auto", StringComparison.OrdinalIgnoreCase))
        {
            var dark = SystemThemeProvider?.IsDarkTheme() ?? true;
            resolved = dark
                ? Wpf.Ui.Appearance.ApplicationTheme.Dark
                : Wpf.Ui.Appearance.ApplicationTheme.Light;
        }
        else if (string.Equals(theme, "Light", StringComparison.OrdinalIgnoreCase))
        {
            resolved = Wpf.Ui.Appearance.ApplicationTheme.Light;
        }
        else
        {
            resolved = Wpf.Ui.Appearance.ApplicationTheme.Dark;
        }
        Wpf.Ui.Appearance.ApplicationThemeManager.Apply(resolved);
    }

    public static void ApplyThemeFromPreferences(string? theme) => ApplyTheme(theme ?? "Dark");

    /// <summary>Resolves the concrete theme name applied for a given preference. "Auto" is resolved via SystemThemeProvider.</summary>
    public static string ResolveActualTheme(string? theme)
    {
        if (string.Equals(theme, "Auto", StringComparison.OrdinalIgnoreCase))
        {
            return (SystemThemeProvider?.IsDarkTheme() ?? true) ? "Dark" : "Light";
        }
        return string.Equals(theme, "Light", StringComparison.OrdinalIgnoreCase) ? "Light" : "Dark";
    }

    [RelayCommand]
    private void ValidateCert()
    {
        var cfg = new CertConfig(
            CertAppId ?? string.Empty,
            CertTenantId ?? string.Empty,
            CertOrganization ?? string.Empty,
            CertThumbprint ?? string.Empty);
        var r = _certValidator.Validate(cfg);
        CertStatusMessage = r.Message;
        CertIsValid = r.IsValid;
    }
}
